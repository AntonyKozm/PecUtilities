using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Xceed.Words.NET;
using System.IO;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data;
using System.Text.RegularExpressions;

namespace PecUtilities
{
    public partial class Converter : UserControl
    {
        private string _tempPath = Directory.GetCurrentDirectory() + @"\temp";
        private string _dbfFilePath;
        private string _dbfFileName;

        public bool Expects = false;

        private readonly ToastViewModel _vm;

        private string _currentPath;

        public Converter() { }

        public Converter(ToastViewModel vm)
        {
            InitializeComponent();
            CheckTemp();

            _vm = vm;
        }

        void ShowMessage(Action<string> action, string text)
        {
            action(text);
        }

        private void CheckTemp()
        {
            Directory.Delete(_tempPath, true);
            Directory.CreateDirectory(_tempPath);
        }

        private void btSearchFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "prz files (*.prz) | *.prz";
            ofd.ShowDialog();
            if (ofd.FileName.Length != 0)
                _currentPath = ofd.FileName;

            FormPaths();
        }

        private void FormPaths()
        {
            string filePath = _currentPath;
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            _dbfFileName = fileName + ".dbf";
            _dbfFilePath = Path.Combine(_tempPath, _dbfFileName);
            if (filePath.Length != 0)
            {
                if (File.Exists(filePath))
                {
                    File.Copy(filePath, _dbfFilePath);
                    Convert();
                }
                else
                {
                    ShowMessage(_vm.ShowError, "Ты пытаешься меня наебать? В смысле файла нет?");
                }
            }
            else
            {
                ShowMessage(_vm.ShowError, "Нет файла - нет результата");
            }
        }

        public void Convert()
        {
            string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _tempPath +
                            ";Extended Properties=dBASE IV;User ID=Admin;Password=;";
            using (OleDbConnection con = new OleDbConnection(constr))
            {
                var sql = "SELECT " +
                          "' ' as '№ пп', " +
                          "P005 as 'Фамилия', " +
                          "P006 as 'Имя', " +
                          "P007 as 'Отчество', " +
                          "K001 as 'Дата рождения', " +
                          "UCase(P012) as 'Место рождения', " +
                          "'Мужской' as 'Пол', " +
                          "'Ю-123456' as 'Личный номер', " +
                          "' ' as 'Обр вид', " +
                          "' ' as 'Обр спец', " +
                          "UCase(FACT_ADDR) as 'Адр регистрации', " +
                          "UCase(FACT_ADDR) as 'Адр проживания', " +
                          "' ' as 'Телефон', " +
                          "P008 as 'Серия', " +
                          "P099 as 'Номер', " +
                          "P010 as 'Дата', " +
                          "UCase(P022) as 'Кем выдан', " +
                          "' ' as 'ВБ Серия', " +
                          "' ' as 'ВБ Номер', " +
                          "' ' as 'ВБ Дата', " +
                          "' ' as 'ВБ кто выдал', " +
                          "' ' as 'СНИЛС', " +
                          "' ' as 'ИНН' " +
                          "FROM " + _dbfFileName +
                          " ORDER BY P005, P006, P007";

                OleDbCommand cmd = new OleDbCommand(sql, con);

                con.Open();
                DataSet ds = new DataSet();
                var da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                con.Close();

                System.Globalization.TextInfo ti = new System.Globalization.CultureInfo("ru-RU", false).TextInfo;

                foreach (DataRow r in ds.Tables[0].Rows)
                    foreach (DataColumn c in ds.Tables[0].Columns)
                    {
                        if (r.IsNull(c) || String.IsNullOrEmpty(r[c].ToString()))
                        {
                            if (c.ColumnName == "'Адр регистрации'" || c.ColumnName == "'Адр проживания'")
                                r[c] = "Нет данных";
                        }

                        if (c.ColumnName == "'Фамилия'" || c.ColumnName == "'Имя'" || c.ColumnName == "'Отчество'")
                            r[c] = ti.ToTitleCase(r[c].ToString().ToLower());

                        if (c.ColumnName == "'Кем выдан'")
                            r[c] = Normalize(r[c].ToString());
                    }

                List<string> columns = new List<string>()
                {
                    "№ пп", "Фамилия", "Имя", "Отчество",
                    "Дата рождения", "Место рождения", "Пол",
                    "Личный номер", "Обр вид", "Обр спец",
                    "Адр регистр", "Адр проживания", "Телефон",
                    "Серия", "Номер", "Дата", "Кем выдан",
                    "ВБ Серия", "ВБ Номер", "ВБ дата",
                    "ВБ кто выдал", "СНИЛС", "ИНН"
                };

                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add("Лист1");
                var worksheet = wb.Worksheets.Worksheet("Лист1");

                worksheet.Cell("C1").Value = "В О";
                worksheet.Cell("C2").Value = "Военный комиссариат Архангельской области г.Архангельск";
                worksheet.Cell("C3").Value = "№ В/Ч";
                worksheet.Cell("E2").Value = GetMilitaryNumber();

                for (int i = 0; i < columns.Count; i++)
                    worksheet.Cell(6, i + 1).Value = columns[i];

                var cell = worksheet.Cell(7, 1);
                cell.Value = ds.Tables[0];

                wb.SaveAs(Path.Combine(Path.GetDirectoryName(_currentPath), GetMilitaryNumber() + ".xlsx"));
                CreateNameList(Path.GetDirectoryName(_currentPath), ds);

                _currentPath = "";
                CheckTemp();

                ShowMessage(_vm.ShowSuccess, "Ковертировано успешно!");
            }
        }

        private string Normalize(string cell)
        {
            var templates = new Dictionary<string, string>()
            {
                { @"\sАО\s", " АРХАНГЕЛЬСКОЙ ОБЛАСТИ " },
                { @"\sНАО\s", " НЕНЕЦКОМУ АВТОНОМНОМУ ОКРУГУ " },
                { @"[\s^]?ОУФМС\s", "ОТДЕЛЕНИЕМ УФМС " },
                { @"[\s^]?МО[\s\d]", "МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ " },
                { @"\sРОСИИ\s", " РОССИИ " },
                { @"\sРОССИ\s", " РОССИИ " },
                { @"[\s^]?ТП\s", "ТЕРРИТОРИАЛЬНЫМ ПУНКТОМ " }
            };

            foreach (var template in templates)
            {
                cell = Regex.Replace(cell.ToString(), template.Key, template.Value, RegexOptions.IgnoreCase);
            }

            return cell.ToUpper();
        }

        private void CreateNameList(string selectedPath, DataSet ds)
        {
            var dataSetTable = ds.Tables[0];
            var rowsCount = dataSetTable.Rows.Count;

            DocX document = DocX.Create(Path.Combine(selectedPath, "Список призывников.docx"));
            var docTable = document.AddTable(rowsCount, 1);
            docTable.Design = Xceed.Document.NET.TableDesign.TableGrid;

            for (int i = 0; i < rowsCount; i++)
            {
                var row = dataSetTable.Rows[i].ItemArray;
                string fullname = String.Format("{0} {1}. {2}.", row[1], 
                                                                 row[2].ToString()[0], 
                                                                 row[3].ToString()[0]);
                docTable.Rows[i]
                        .Cells[0]
                        .Paragraphs[0].Append(fullname)
                                      .Font("Times New Roman");
            }

            document.InsertParagraph()
                    .InsertTableAfterSelf(docTable);

            document.Save();
        }

        private string GetMilitaryNumber()
        {
            Dictionary<string, string> militaryNumber = new Dictionary<string, string>()
            {
                { "08413500", "Арх-" },
                { "08414391", "Севск-" },
                { "08413724", "Вельск-" },
                { "08413782", "Вилег-" },
                { "08414008", "Коноша-" },
                { "08413888", "Котлас-" },
                { "08414014", "Красноб-" },
                { "08414215", "Мезень-" },
                { "08414221", "НАО-" },
                { "08414289", "Нянд-" },
                { "08414327", "Онега-" },
                { "08414385", "Пинега-" },
                { "08414333", "Плес-" },
                { "08414343", "Прим-" },
                { "08414497", "Холм-" },
            };

            string date = DateTime.Now.Date.ToString("dd.MM");
            return militaryNumber[Path.GetFileNameWithoutExtension(_dbfFileName)] + date;
        }

        private void UserControl_Drop(object sender, DragEventArgs e)
        {
            if (e.Data is DataObject && ((DataObject)e.Data).ContainsFileDropList())
            {
                var dropList = ((DataObject)e.Data).GetFileDropList();
                
                if (dropList.Count == 1)
                {
                    foreach (string filePath in dropList)
                    {
                        if (Path.HasExtension(filePath))
                        {
                            if (filePath.Substring(filePath.LastIndexOf('.')) == ".prz")
                            {
                                _currentPath = filePath;
                                FormPaths();
                                _currentPath = "";
                            }
                            else
                            {
                                ShowMessage(_vm.ShowError, "Файл не того формата!");
                                return;
                            }
                        }
                        else
                        {
                            ShowMessage(_vm.ShowError, "Папка не подойдёт!");
                            return;
                        }
                    }
                }
                else
                {
                    ShowMessage(_vm.ShowInformation, "Перетащил больше 1 файла, на исходную...");
                    return;
                }
            }
        }

        public void SetDragMaskIndex(int index)
        {
            Panel.SetZIndex(dropEffect, index);
        }

        private void main_Drop(object sender, DragEventArgs e)
        {
            SetDragMaskIndex(-1);
        }
    }
}

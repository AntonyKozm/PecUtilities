using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Xceed.Words.NET;
using Ookii.Dialogs.Wpf;

namespace PecUtilities
{
    /// <summary>
    /// Логика взаимодействия для MagazineCreator.xaml
    /// </summary>
    public partial class MagazineCreator : UserControl
    {
        private string _templatesPath = Path.Combine(Directory.GetCurrentDirectory(), @"templates");

        private readonly ToastViewModel _vm;

        public MagazineCreator() { }

        public MagazineCreator(ToastViewModel vm)
        {
            InitializeComponent();

            _vm = vm;
        }

        void ShowMessage(Action<string> action, string text)
        {
            action(text);
        }

        private void btSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog();
            dialog.Description = "Выберите папку";
            dialog.UseDescriptionForTitle = true;

            if ((bool)dialog.ShowDialog(Window.GetWindow(this)))
            {
                try
                {
                    var names = ReadNameLists(dialog.SelectedPath);

                    CreateLaundryMagazine(names, dialog.SelectedPath);
                    CreateThermometryMagazine(names, dialog.SelectedPath);

                    ShowMessage(_vm.ShowSuccess, "Журналы созданы!");
                }
                catch (Exception ex)
                {
                    ShowMessage(_vm.ShowError, "Ошибка создания журналов!\n\n" + ex.Message);
                    return;
                }
            }
        }

        private List<string> ReadNameLists(string selectedPath)
        {
            string[] allFoundFiles = Directory.GetFiles(selectedPath, "Список призывников.docx", SearchOption.AllDirectories);
            var names = new List<string>();

            foreach (var filepath in allFoundFiles)
            {
                var doc = DocX.Load(filepath);

                foreach (var row in doc.Tables[0].Rows)
                    names.Add(row.Cells[0].Xml.Value);
            }

            return names;
        }

        private string CopyMagazine(string destinationPath, string magazineName)
        {
            string templatePath = Path.Combine(_templatesPath, magazineName);
            string savePath = Path.Combine(destinationPath, magazineName);

            if (File.Exists(savePath))
                File.Delete(savePath);

            File.Copy(templatePath, savePath);

            return savePath;
        }

        private void CreateLaundryMagazine(List<string> names, string destinationPath)
        {
            string savePath = CopyMagazine(destinationPath, "Журнал выдачи белья.docx");

            var magazine = DocX.Load(savePath);
            var tables = magazine.Tables;

            for (int i = 0; i < names.Count; i++)
            {
                int index = (int)Math.Floor((decimal)i / 46);
                tables[index].Rows[i + 1 - 46 * index]
                             .Cells[1]
                             .Paragraphs[0].Append(names[i])
                                           .Font("Times New Roman")
                                           .FontSize(11);
            }

            magazine.Save();
        }

        private void CreateThermometryMagazine(List<string> names, string destinationPath)
        {
            string savePath = CopyMagazine(destinationPath, "Журнал термометрии.docx");

            var magazine = DocX.Load(savePath);
            var tables = magazine.Tables;

            for (int i = 0; i < names.Count; i++)
            {
                int index = (int)Math.Floor((decimal)i / 90);
                int half = (int)Math.Floor((decimal)i / 45) % 2;
                int rowIndex = i + 1 - 45 * half - 90 * index;
                int columnIndex = 1;

                if (half == 1)
                    columnIndex = 4;

                tables[index].Rows[rowIndex]
                             .Cells[columnIndex]
                             .Paragraphs[0].Append(names[i])
                                           .Font("Times New Roman")
                                           .FontSize(11);
            }

            magazine.Save();
        }
    }
}

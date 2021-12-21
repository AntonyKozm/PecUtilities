using Ionic.Zip;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace PecUtilities
{
    /// <summary>
    /// Логика взаимодействия для DeleteNumbers.xaml
    /// </summary>
    public partial class DeleteNumbers : UserControl
    {
        private string _tempPath = Directory.GetCurrentDirectory() + @"\temp";
        private List<string> _tempFolders = new List<string>();

        public DeleteNumbers()
        {
            InitializeComponent();
        }

        private void CheckTemp()
        {
            Directory.Delete(_tempPath, true);
            Directory.CreateDirectory(_tempPath);
        }

        private void bField_Drop(object sender, DragEventArgs e)
        {
            List<string> files = new List<string>();

            if (e.Data is DataObject && ((DataObject)e.Data).ContainsFileDropList())
                foreach (string filePath in ((DataObject)e.Data).GetFileDropList())
                    files.Add(filePath);

            CheckTemp();
            Extract(files);
        }

        private void Extract(List<string> files)
        {
            foreach (string path in files)
            {
                string pathInTemp = Path.Combine(_tempPath, Path.GetFileNameWithoutExtension(path));

                if (Directory.Exists(pathInTemp))
                    Directory.Delete(pathInTemp);
                Directory.CreateDirectory(pathInTemp);

                _tempFolders.Add(pathInTemp);

                var options = new ReadOptions();
                options.Encoding = Encoding.GetEncoding(866);

                using (ZipFile zip = ZipFile.Read(path, options))
                {
                    foreach (ZipEntry e in zip)
                        e.Extract(pathInTemp);
                }
            }

            DeleteNum();
        }

        private void DeleteNum()
        {
            foreach (var folder in _tempFolders)
            {
                var names = new List<string>();
                foreach (var file in Directory.GetFiles(folder, "*.*")
                                              .Where(s => s.EndsWith(".piu")))
                {
                    var piuName = Path.GetFileName(file);
                    var jpgName = Path.GetFileNameWithoutExtension(piuName) + ".jpg";

                    var splittedName = piuName.Split('.')[0].Split('-');
                    var length = splittedName.Length;

                    var newName = splittedName[length - 2] + "-" + splittedName[length - 1];

                    names.Add(newName);

                    if (names.FindAll(s => s == newName).Count > 1)
                        newName += "-" + (names.FindAll(s => s == newName).Count - 1).ToString();

                    var jpgNewName = newName + ".jpg";
                    newName += ".piu";

                    File.Move(Path.Combine(folder, piuName), Path.Combine(folder, newName));
                    File.Move(Path.Combine(folder, jpgName), Path.Combine(folder, jpgNewName));
                }
            }

            Process.Start(_tempPath);
        }
    }
}

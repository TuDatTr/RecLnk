using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace RecLnk
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        int maxFileCount = 0;
        readonly Stopwatch s = new Stopwatch();
        IEnumerable<string> fileList;
        public static void CreateShortcut(string sourceFileLocation, string lnkPath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkPath);

            shortcut.Description = "Created by RecLnk";   // The description of the shortcut
            // shortcut.IconLocation = srcLocation;           // The icon of the shortcut
            shortcut.TargetPath = sourceFileLocation;                 // The path of the file that will launch when the shortcut is run
            shortcut.Save();                                    // Save the shortcut
        }

        public IEnumerable<string> Shuffle(IEnumerable<string> list)
        {
            var r = new Random();
            var shuffledList =
                list.
                    Select(x => new { Number = r.Next(), Item = x }).
                    OrderBy(x => x.Number).
                    Select(x => x.Item).
                    Take(maxFileCount-1); // Assume first @size items is fine
            return shuffledList.ToList();
        }

        private void ChooseFolder()
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                InitialDirectory = tbTarget.Text, // Use current value for initial dir
                Title = "Select a Directory", // instead of default "Save As"
                Filter = "Directory|*.this.directory", // Prevents displaying files
                FileName = "select" // Filename will then be "select.this.directory"
            };
            Microsoft.Win32.SaveFileDialog dialog = saveFileDialog;
            if (dialog.ShowDialog() == true)
            {
                string path = dialog.FileName;
                // Remove fake filename from resulting path
                path = path.Replace("\\select.this.directory", "");
                path = path.Replace(".this.directory", "");
                // If user has changed the filename, create the new directory
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                // Our final value is in path
                tbTarget.Text = path;
            }
        }

        private void BtnDirChooser_Click(object sender, RoutedEventArgs e)
        {
            ChooseFolder();
            var dirPath = tbTarget.Text;
            fileList = GetRecursiveFiles(dirPath);
        }

        private IEnumerable<string> GetRecursiveFiles(string dirPath)
        {
            try
            {
                var fileList = Directory.EnumerateFiles(dirPath, "*.*", SearchOption.AllDirectories);
                int fileCount = 0;
                foreach (var file in fileList)
                {
                    fileCount++;
                }
                SetProgress("Now press the \"Go!\" button.");
                if (fileList != null)
                {
                    tbCount.Text = fileCount.ToString();
                    maxFileCount = fileCount;
                }
                return fileList;
            } catch (Exception)
            {
                MessageBox.Show("Please enter a valid path");
                return null;
            }
        }

        private void SetProgress(string prog)
        {
            lblProgress.Text = prog;
        }
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{
                SetProgress("Running...");
                int sampleSize = Convert.ToInt32(tbCount.Text);
                if (sampleSize <= maxFileCount)
                {
                    var dirPath = AppDomain.CurrentDomain.BaseDirectory;
                    var date = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                    var targetDir = System.IO.Path.Combine(dirPath, new DirectoryInfo(tbTarget.Text).Name, date);
                    fileList = Shuffle(fileList);
                    System.IO.Directory.CreateDirectory(targetDir);
                    for (int i = 0; i < sampleSize; i++)
                    {
                        string lnkPath = System.IO.Path.Combine(targetDir, System.IO.Path.GetFileName(fileList.ElementAt(i)) + ".lnk");
                        CreateShortcut(fileList.ElementAt(i), lnkPath);
                    }
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        Arguments = targetDir,
                        FileName = "explorer.exe"
                    };
                    Process.Start(startInfo);
                    SetProgress("Done.");
                }
                else
                {
                    MessageBox.Show("Sample size too big.\nMax sample size: " + maxFileCount);
                    SetProgress("Now press the \"Go!\" button.");
                }
            //}catch (Exception)
            //{
            //    MessageBox.Show("Please enter a valid sample size.\nMax sample size: " + maxFileCount);
            //    setProgress("Now press the \"Go!\" button.");
            //}
            
        }

        private void TbTarget_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SetProgress("Scanning files...");
                var dirPath = tbTarget.Text;
                fileList = GetRecursiveFiles(dirPath);
            }
        }

        private void TbTarget_GotFocus(object sender, RoutedEventArgs e)
        {
            if (tbTarget.Text == "Directory...")
                tbTarget.Text = "";
        }

        private void FormMain_Initialized(object sender, EventArgs e)
        {
            SetProgress("Ready...");
        }
    }
}

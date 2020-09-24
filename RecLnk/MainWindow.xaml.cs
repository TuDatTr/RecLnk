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
        public static void CreateShortcut(string sourceFileLocation, string lnkPath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkPath);

            shortcut.Description = "Created by RecLnk";   // The description of the shortcut
            // shortcut.IconLocation = srcLocation;           // The icon of the shortcut
            shortcut.TargetPath = sourceFileLocation;                 // The path of the file that will launch when the shortcut is run
            shortcut.Save();                                    // Save the shortcut
        }
        public List<string> Shuffle(List<string> list)
        {
            Random rng = new Random();
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                string value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
            return list;
        }

        private void ChooseFolder()
        {
            setProgress("Scanning files...");
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

        private void btnDirChooser_Click(object sender, RoutedEventArgs e)
        {
            ChooseFolder();
            var dirPath = tbTarget.Text;
            var fileList = getRecursiveFiles(dirPath);
            if (fileList != null)
            {
                if (fileList.Count != 0)
                {
                    tbCount.Text = fileList.Count.ToString();
                    maxFileCount = fileList.Count;
                }
            }
        }

        private List<string> getRecursiveFiles(string dirPath)
        {
            try
            {
                List<string> fileList = new List<string>();
                string[] files = Directory.GetFiles(dirPath, "*.*", SearchOption.AllDirectories);
                tbLog.Clear();
                foreach (var file in files)
                {
                    tbLog.Text += file + "\r\n";
                }
                setProgress("Now press the \"Go!\" button.");
                return files.ToList();
            } catch (Exception)
            {
                MessageBox.Show("Please enter a valid path");
                return null;
            }
        }

        private void setProgress(string prog)
        {
            lblProgress.Text = prog;
        }
        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                setProgress("Running...");
                int sampleSize = Convert.ToInt32(tbCount.Text);
                if (sampleSize <= maxFileCount)
                {
                    var dirPath = AppDomain.CurrentDomain.BaseDirectory;
                    var date = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                    var targetDir = System.IO.Path.Combine(dirPath, new DirectoryInfo(tbTarget.Text).Name, date);
                    var fileList = tbLog.Text.Split("\r\n").ToList();
                    fileList = Shuffle(fileList);
                    string lnkPath = "";
                    System.IO.Directory.CreateDirectory(targetDir);
                    for (int i = 0; i < sampleSize; i++)
                    {
                        lnkPath = System.IO.Path.Combine(targetDir, System.IO.Path.GetFileName(fileList[i])) + ".lnk";
                        CreateShortcut(fileList[i], lnkPath);
                    }
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        Arguments = targetDir,
                        FileName = "explorer.exe"
                    };
                    Process.Start(startInfo);
                    setProgress("Done.");
                }
                else
                {
                    MessageBox.Show("Sample size too big.\nMax sample size: " + maxFileCount);
                }
            }catch (Exception)
            {
                MessageBox.Show("Please enter a valid sample size.\nMax sample size: " + maxFileCount);
            }
            
        }

        private void tbTarget_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                setProgress("Scanning files...");
                var dirPath = tbTarget.Text;
                var fileList = getRecursiveFiles(dirPath);
                if (fileList != null)
                {
                    if (fileList.Count != 0)
                    {
                        tbCount.Text = fileList.Count.ToString();
                        maxFileCount = fileList.Count;
                    }
                }
            }
        }

        private void tbTarget_GotFocus(object sender, RoutedEventArgs e)
        {
            if (tbTarget.Text == "Directory...")
                tbTarget.Text = "";
        }

        private void FormMain_Initialized(object sender, EventArgs e)
        {
            setProgress("Ready...");
        }
    }
}

using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
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
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.InitialDirectory = tbTarget.Text; // Use current value for initial dir
            dialog.Title = "Select a Directory"; // instead of default "Save As"
            dialog.Filter = "Directory|*.this.directory"; // Prevents displaying files
            dialog.FileName = "select"; // Filename will then be "select.this.directory"
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
                return files.ToList();
            } catch (Exception)
            {
                MessageBox.Show("Please enter a valid path");
                return null;
            }
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int sampleSize = Convert.ToInt32(tbCount.Text);
                if (sampleSize <= maxFileCount)
                {
                    var dirPath = tbTarget.Text;
                    var fileList = tbLog.Text.Split("\r\n").ToList();
                    fileList = Shuffle(fileList);
                    string lnkPath = "";
                    for (int i = 0; i < sampleSize; i++)
                    {
                        lnkPath = System.IO.Path.Combine(dirPath, System.IO.Path.GetFileName(fileList[i])) + ".lnk";
                        CreateShortcut(fileList[i], lnkPath);
                    }
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
    }
}

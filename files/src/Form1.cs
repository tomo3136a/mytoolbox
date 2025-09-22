using System;
using System.IO;
using System.Windows.Forms;

namespace files
{
    public partial class Form1 : Form
    {
        const string AppName = "files";

        public Form1()
        {
            InitializeComponent();
        }

        void OnSelect(Object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.SelectedPath = Directory.GetCurrentDirectory();
            if (DialogResult.OK == dlg.ShowDialog())
            {
                this.txt_1.Text = dlg.SelectedPath;
                //this.txt_1.BackColor = Color.White;
                //this.txt_1.BackColor = SystemColors.Window;
            }
            dlg.Dispose();
        }

        void OnClose(Object sender, EventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                var s = this.txt_1.Text;
                if (!Directory.Exists(s))
                {
                    this.txt_1.Focus();
                    //this.txt_1.BackColor = Color.Yellow;
                    MessageBox.Show(this, "フォルダを選択してください。", AppName);
                    this.txt_1.SelectAll();
                    return;
                }
                if (FileList())
                {
                    MessageBox.Show(this, "Compleated.", AppName);
                }
            }
            this.Close();
        }

        bool FileList()
        {
            var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var dt = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            var op = Path.Combine(doc, "dirlist_" + dt + ".txt");
            if (File.Exists(op))
            {
                File.Delete(op);
            }
            var p = txt_1.Text;
            if (!Directory.Exists(p))
            {
                return false;
            }
            using (StreamWriter ost = new StreamWriter(op, true))
            {
                var line = "path: " + p;
                ost.WriteLine(line);
                WriteFileList(ost, p);
            }
            return true;
        }

        static void WriteFileList(StreamWriter ost, string path)
        {
            DirectoryInfo diTop = new DirectoryInfo(path);
            int sz = diTop.FullName.Length + 1;
            try
            {
                foreach (var fi in diTop.EnumerateFiles())
                {
                    try
                    {
                        var line = fi.FullName + "\t\t" + fi.Length.ToString("N0");
                        ost.WriteLine(line.Substring(sz));
                    }
                    catch (UnauthorizedAccessException unAuthTop)
                    {
                        Console.WriteLine("" + unAuthTop.Message);
                    }
                }

                foreach (var di in diTop.EnumerateDirectories("*"))
                {
                    try
                    {
                        foreach (var fi in di.EnumerateFiles("*", SearchOption.AllDirectories))
                        {
                            try
                            {
                                var line = fi.FullName + "\t\t" + fi.Length.ToString("N0");
                                ost.WriteLine(line.Substring(sz));
                            }
                            catch (UnauthorizedAccessException unAuthFile)
                            {
                                Console.WriteLine("unAuthFile: " + unAuthFile.Message);
                            }
                        }
                    }
                    catch (UnauthorizedAccessException unAuthSubDir)
                    {
                        Console.WriteLine("unAuthSubDir: " + unAuthSubDir.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException dirNotFound)
            {
                Console.WriteLine("" + dirNotFound.Message);
            }
            catch (UnauthorizedAccessException unAuthDir)
            {
                Console.WriteLine("unAuthDir: " + unAuthDir.Message);
            }
            catch (PathTooLongException longPath)
            {
                Console.WriteLine("" + longPath.Message);
            }
        }
    }
}

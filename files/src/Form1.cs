using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace files
{
    public partial class Form1 : Form
    {
        const string AppName = "files";
        private long ticks = 0;

        public Form1()
        {
            InitializeComponent();
        }

        void OnSelect(Object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            var root = txt_1.Text;
            if (Directory.Exists(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            dlg.SelectedPath = root;
            if (DialogResult.OK == dlg.ShowDialog())
            {
                this.txt_1.Text = dlg.SelectedPath;
                this.txt_1.BackColor = Color.White;
                this.txt_1.BackColor = SystemColors.Window;
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
                    this.txt_1.BackColor = Color.Yellow;
                    MessageBox.Show(this, "フォルダを選択してください。", AppName);
                    this.txt_1.SelectAll();
                    return;
                }

                lbl_2.Text = "処理中";
                this.Update();
                var mode = this.rb_1.Checked ? 1 : rb_2.Checked ? 2 : 3;
                var bTree = this.cb_1.Checked;
                var bSize = this.cb_2.Checked;
                var bDate = this.cb_3.Checked;
                if (FileList(this.txt_1.Text, mode, bTree, bSize, bDate))
                {
                    string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                    File.AppendAllLines(AppName + ".log", msg.Split('\n'));
                    lbl_2.Text = msg;
                    this.Update();
                    MessageBox.Show(this, "Successfull.", AppName);
                }
                else
                {
                    string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                    lbl_2.Text = msg;
                    this.Update();
                    MessageBox.Show(this, "Denied.", AppName);
                }
            }
            this.Close();
        }

        bool FileList(string p, int mode, bool bTree, bool bSize, bool bDate)
        {
            var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var dt = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            var op = Path.Combine(doc, "dirlist_" + dt + ".txt");
            if (File.Exists(op))
            {
                File.Delete(op);
            }
            if (!Directory.Exists(p))
            {
                return false;
            }
            long StartTicks = DateTime.Now.Ticks;
            using (StreamWriter ost = new StreamWriter(op, true))
            {
                var line = "path: " + p;
                ost.WriteLine(line);
                switch (mode)
                {
                    case 1: // file list
                        WriteFileList(ost, p, bTree, bSize, bDate);
                        break;
                    case 2: // directory list
                        WriteDirList(ost, p, bTree, bSize, bDate);
                        break;
                    case 3: // file+directory list
                        WriteFileDirList(ost, p, bTree, bSize, bDate);
                        break;
                    default:
                        break;
                }
            }
            ticks = DateTime.Now.Ticks - StartTicks;
            return true;
        }

        static void WriteFileList(StreamWriter ost, string path, bool bTree, bool bSize, bool bDate)
        {
            DirectoryInfo root = new DirectoryInfo(path);
            string p = root.FullName;
            int sz = root.FullName.Length;
            if (p[sz - 1] != '\\' && p[sz - 1] != '/') sz += 1;
            Console.WriteLine("> " + sz + " " + root);
            try
            {
                foreach (var fi in root.EnumerateFiles())
                {
                    try
                    {
                        p = fi.FullName;
                        var line = "\t" + fi.FullName.Substring(sz);
                        Console.WriteLine("> " + sz + " " + p + line);
                        if (bSize) line += "\t" + fi.Length.ToString("N0");
                        if (bDate) line += "\t" + fi.LastWriteTime.ToLocalTime();
                        ost.WriteLine(line);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("# UnauthorizedAccess[1]: " + ex.Message);
                    }
                }

                foreach (var di in root.EnumerateDirectories("*"))
                {
                    try
                    {
                        foreach (var fi in di.EnumerateFiles("*", SearchOption.AllDirectories))
                        {
                            try
                            {
                                var line = "\t" + fi.FullName.Substring(sz);
                                if (bSize) line += "\t" + fi.Length.ToString("N0");
                                if (bDate) line += "\t" + fi.LastWriteTime.ToLocalTime();
                                ost.WriteLine(line);
                            }
                            catch (UnauthorizedAccessException ex)
                            {
                                ost.WriteLine("# UnauthorizedAccess[2]: " + ex.Message);
                            }
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("# UnauthorizedAccess[3]: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                ost.WriteLine("# DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ost.WriteLine("# UnauthorizedAccess[4]: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                ost.WriteLine("# PathTooLong: " + ex.Message);
            }
            catch (Exception ex)
            {
                ost.WriteLine("# Exception: " + ex.Message);
            }
        }

        public static IEnumerable<DirectoryInfo> EnumerateDirectories(DirectoryInfo di)
        {
            var dirs = Enumerable.Empty<DirectoryInfo>();
            try
            {
                dirs = di.EnumerateDirectories().Where(
                    v => (v.Attributes & FileAttributes.System) != FileAttributes.System);
                dirs = dirs
                        .Aggregate<DirectoryInfo, IEnumerable<DirectoryInfo>>(
                            dirs,
                            (a, v) => a.Union(EnumerateDirectories(v))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            return dirs;
        }

        static void WriteDirList(StreamWriter ost, string path, bool bTree, bool bSize, bool bDate)
        {
            DirectoryInfo root = new DirectoryInfo(path);
            string p = root.FullName;
            int sz = p.Length;
            if (p[sz - 1] != '\\' && p[sz - 1] != '/') sz += 1;
            try
            {
                foreach (var di in EnumerateDirectories(root))
                {
                    try
                    {
                        var line = "\t" + di.FullName.Substring(sz);
                        if (bSize) line += "\t";
                        if (bDate) line += "\t" + di.LastWriteTime.ToLocalTime();
                        ost.WriteLine(line);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("# UnauthorizedAccess[2]: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                ost.WriteLine("# DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ost.WriteLine("# UnauthorizedAccess[4]: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                ost.WriteLine("# PathTooLong: " + ex.Message);
            }
            catch (Exception ex)
            {
                ost.WriteLine("# Exception: " + ex.Message);
            }
        }

        static void WriteFileDirList(StreamWriter ost, string path, bool bTree, bool bSize, bool bDate)
        {
            DirectoryInfo root = new DirectoryInfo(path);
            string p = root.FullName;
            int sz = p.Length;
            if (p[sz - 1] != '\\' && p[sz - 1] != '/') sz += 1;
            try
            {
                foreach (var di in EnumerateDirectories(root))
                {
                    try
                    {
                        var line = "\t" + di.FullName.Substring(sz) + "\\";
                        if (bSize) line += "\t";
                        if (bDate) line += "\t" + di.LastWriteTime.ToLocalTime();
                        ost.WriteLine(line);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("# UnauthorizedAccess[1]: " + ex.Message);
                    }
                    try
                    {
                        foreach (var fi in di.EnumerateFiles())
                        {
                            try
                            {
                                var line = "\t" + fi.FullName.Substring(sz);
                                if (bSize) line += "\t" + fi.Length.ToString("N0");
                                if (bDate) line += "\t" + fi.LastWriteTime.ToLocalTime();
                                ost.WriteLine(line);
                            }
                            catch (UnauthorizedAccessException ex)
                            {
                                ost.WriteLine("# UnauthorizedAccess[2]: " + ex.Message);
                            }
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("# UnauthorizedAccess[3]: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                ost.WriteLine("# DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ost.WriteLine("# UnauthorizedAccess[4]: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                ost.WriteLine("# PathTooLong: " + ex.Message);
            }
            catch (Exception ex)
            {
                ost.WriteLine("# Exception: " + ex.Message);
            }
        }
    }
}

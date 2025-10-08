using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Security.Cryptography.Xml;

namespace files
{
    public partial class Form1 : Form
    {
        private static bool l_bSize;
        private static bool l_bDate;
        private static bool l_bFile;
        private static List<string> err_list = new List<string>();

        static void SetMode(bool bSize, bool bDate, bool bFile)
        {
            l_bSize = bSize;
            l_bDate = bDate;
            l_bFile = bFile;
        }

        public static IEnumerable<DirectoryInfo> EnumerateDirectories(DirectoryInfo di)
        {
            var m = FileAttributes.Hidden | FileAttributes.System;
            IEnumerable<DirectoryInfo> dia = null;
            var dcnt = 0;
            try
            {
                dia = di.EnumerateDirectories();
                dcnt = dia.Count();
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            if (dcnt > 0)
            {
                foreach (var a in dia)
                {
                    if (a == null)
                    {
                        err_list.Add("!!![1-2] a is null.");
                        continue;
                    }
                    if ((a.Attributes & m) == 0)
                    {
                        yield return a;
                        IEnumerable<DirectoryInfo> dib = null;
                        try
                        {
                            dib = EnumerateDirectories(a);
                        }
                        catch (System.UnauthorizedAccessException)
                        {
                        }
                        if (dib == null)
                        {
                            continue;
                        }
                        foreach (var b in dib)
                        {
                            yield return b;
                        }
                    }
                }
            }
        }

        public static IEnumerable<string> EnumerateDirectoriesTree(DirectoryInfo di, string tab = "", bool bEnd = true)
        {
            var m = FileAttributes.Hidden | FileAttributes.System;
            IEnumerable<DirectoryInfo> dia = null;
            var dcnt = 0;
            try
            {
                dia = di.EnumerateDirectories().Where((v) => (v.Attributes & m) == 0);
                dcnt = dia.Count();
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            if (l_bFile)
            {
                IEnumerable<FileInfo> fia = null;
                var fcnt = 0;
                try
                {
                    fia = di.EnumerateFiles().Where((v) => (v.Attributes & m) == 0);
                    fcnt = fia.Count();
                }
                catch (System.UnauthorizedAccessException)
                {
                }
                if (fcnt > 0)
                {
                    var ftab = tab + ((dcnt > 0) ? "│  " : "    ");
                    foreach (var fi in fia)
                    {
                        yield return (ftab + fi.Name);
                    }
                    yield return ftab;
                }
            }
            if (dcnt > 0)
            {
                var idx = 0;
                var cnt = dia.Count();
                foreach (var a in dia)
                {
                    idx++;
                    var btab = tab + ((idx == cnt) ? "└─" : "├─");
                    if (a != null) yield return (btab + a.Name);
                    IEnumerable<string> dib = null;
                    try
                    {
                        var dtab = tab + ((cnt == 0) ? "" : (idx == cnt) ? "    " : "│  ");
                        dib = EnumerateDirectoriesTree(a, dtab, (idx == cnt));
                    }
                    catch (System.UnauthorizedAccessException)
                    {
                    }
                    if (dib != null)
                    {
                        foreach (var b in dib)
                        {
                            yield return b;
                        }
                    }
                }
            }
        }

        bool FileList(string src, string dst, int mode, bool bTree, bool bSize, bool bDate)
        {
            if (!Directory.Exists(src))
            {
                return false;
            }
            var p = (src[src.Length - 1] == '\\') ? src : (src + "\\");
            var root = new DirectoryInfo(p);
            if (File.Exists(dst))
            {
                File.Delete(dst);
            }
            var bFile = (mode != 2);
            var bDir = (mode != 1);
            SetMode(bSize, bDate, bFile);
            long StartTicks = DateTime.Now.Ticks;
            using (StreamWriter ost = new StreamWriter(dst, true))
            {
                var line = Path.GetFileName(src);
                if (line.Length > 0)
                {
                    ost.WriteLine(line);
                }
                if (!bTree)
                {
                    WriteFileDirList(ost, root, bFile, bDir);
                }
                else
                {
                    WriteFileDirTree(ost, root, bFile, bDir);
                }
            }
            ticks = DateTime.Now.Ticks - StartTicks;

            if (err_list.Count() > 0)
            {
                var msg = "";
                foreach (var m in err_list)
                {
                    msg += m + "\n";
                }
                MessageBox.Show(msg, AppName);
            }
            return true;
        }

        static void WriteFileDirList(StreamWriter ost, DirectoryInfo root, bool bFile, bool bDir)
        {
            int sz = root.FullName.Length;
            try
            {
                foreach (var di in EnumerateDirectories(root))
                {
                    try
                    {
                        if (bDir)
                        {
                            var s = di.FullName.Substring(sz);
                            if (s.Length > 0)
                            {
                                var line = "\t" + s + "\\";
                                if (l_bSize) line += "\t";
                                if (l_bDate) line += "\t" + di.LastWriteTime.ToLocalTime();
                                ost.WriteLine(line);
                            }
                        }
                        if (bFile)
                        {
                            IEnumerable<FileInfo> fia = null;
                            try
                            {
                                fia = di.EnumerateFiles();
                            }
                            catch (UnauthorizedAccessException ex)
                            {
                                err_list.Add("!!! UnauthorizedAccess dirctory: " + ex.Message);
                                fia = null;
                            }
                            if (fia != null)
                            {
                                foreach (var fi in fia)
                                {
                                    try
                                    {
                                        var line = "\t" + fi.FullName.Substring(sz);
                                        if (l_bSize) line += "\t" + fi.Length.ToString("N0");
                                        if (l_bDate) line += "\t" + fi.LastWriteTime.ToLocalTime();
                                        ost.WriteLine(line);
                                    }
                                    catch (UnauthorizedAccessException ex)
                                    {
                                        err_list.Add("!!! UnauthorizedAccess file: " + ex.Message);
                                    }
                                }
                            }
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        err_list.Add("!!! UnauthorizedAccess directory: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                err_list.Add("!!! DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                err_list.Add("!!! UnauthorizedAccess: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                err_list.Add("!!! PathTooLong: " + ex.Message);
            }
        }

        static void WriteFileDirTree(StreamWriter ost, DirectoryInfo root, bool bFile, bool bDir)
        {
            try
            {
                foreach (var line in EnumerateDirectoriesTree(root))
                {
                    ost.WriteLine(line);
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                err_list.Add("!!! DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                err_list.Add("!!! UnauthorizedAccess: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                err_list.Add("!!! PathTooLong: " + ex.Message);
            }
        }
    }
}

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
        private static bool l_bSize;
        private static bool l_bDate;

        static void SetMode(bool bSize, bool bDate)
        {
            l_bSize = bSize;
            l_bDate = bDate;
        }

        public static IEnumerable<DirectoryInfo> EnumerateDirectories(DirectoryInfo di)
        {
            var dirs = Enumerable.Empty<DirectoryInfo>();
            try
            {
                dirs = di.EnumerateDirectories()
                    .Where(v => (v.Attributes & FileAttributes.System) == 0)
                    .Aggregate<DirectoryInfo, IEnumerable<DirectoryInfo>>(
                        dirs.Append(di),
                        (a, v) => a.Union(EnumerateDirectories(v))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            return dirs;
        }

        public static IEnumerable<string> EnumerateDirectories2(string dp, ref int idx, int cnt = 0, string tab = "")
        {
            var dirs = Enumerable.Empty<string>();
            try
            {
                var col = Directory.EnumerateDirectories(dp);
                var cnt2 = col.Count();
                var idx2 = 0;
                var tab2 = ((idx + 1) == cnt) ? "    " : "│  ";
                if (cnt == 0) tab2 = "";
                var tab3 = ((idx + 1) == cnt) ? "└─" : "├─";
                var tab4 = (Directory.GetDirectories(dp).Count() == 0) ? "    " : "│  ";
                dirs = col
                    .Aggregate<string, IEnumerable<string>>(
                        dirs.Append(tab + tab3 + Path.GetFileName(dp))
                        .Union(Directory.EnumerateFiles(dp).Select(
                            (v) => tab + tab2 + tab4 + Path.GetFileName(v))),
                        (a, v) => a.Union(EnumerateDirectories2(v, ref idx2, cnt2, tab + tab2))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            idx++;
            if (dirs.Count() > 0)
            {

            }
            return dirs;
        }

        bool FileList(string p, int mode, bool bTree, bool bSize, bool bDate)
        {
            var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var dt = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var op = Path.Combine(doc, "dirlist_" + dt + ".txt");
            if (File.Exists(op))
            {
                File.Delete(op);
            }
            if (!Directory.Exists(p))
            {
                return false;
            }
            SetMode(bSize, bDate);
            long StartTicks = DateTime.Now.Ticks;
            using (StreamWriter ost = new StreamWriter(op, true))
            {
                var line = "path: " + p;
                ost.WriteLine(line);
                var bFile = (mode != 2);
                var bDir = (mode != 1);
                if (!bTree)
                {
                    WriteFileDirList(ost, p, bFile, bDir);
                }
                else
                {
                    WriteFileDirTree(ost, p, bFile, bDir);
                }
            }
            ticks = DateTime.Now.Ticks - StartTicks;
            return true;
        }

        static void WriteFileDirList(StreamWriter ost, string path, bool bFile, bool bDir)
        {
            string p = (path[path.Length - 1] == '\\') ? path : (path + "\\");
            DirectoryInfo root = new DirectoryInfo(p);
            p = root.FullName;
            int sz = p.Length;
            bool bSubDir = false;
            try
            {
                foreach (var di in EnumerateDirectories(root))
                {
                    try
                    {
                        if (bSubDir)
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
                        bSubDir = bDir;
                        if (bFile)
                        {
                            foreach (var fi in di.EnumerateFiles())
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
                                    ost.WriteLine("!!! UnauthorizedAccess file: " + ex.Message);
                                }
                            }
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("!!! UnauthorizedAccess directory: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                ost.WriteLine("!!! DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ost.WriteLine("!!! UnauthorizedAccess: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                ost.WriteLine("!!! PathTooLong: " + ex.Message);
            }
            catch (Exception ex)
            {
                ost.WriteLine("!!! Exception: " + ex.Message);
            }
        }

        static void WriteFileDirTree(StreamWriter ost, string path, bool bFile, bool bDir)
        {
            string p = (path[path.Length - 1] == '\\') ? path : (path + "\\");
            int sz = p.Length;
            int cnt = p.Count((c) => c == '\\');
            int cnt1 = cnt;
            bool bSubDir = false;
            try
            {
                var idx = 0;
                foreach (var dp in EnumerateDirectories2(p, ref idx))
                {
                    try
                    {
                        if (bSubDir)
                        {
                            var line = dp;
                            ost.WriteLine(line);
                        }
                        bSubDir = bDir;
                        if (bFile)
                        {
                            var ftab = "    ";
                            foreach (var fp in Directory.EnumerateFiles(dp))
                            {
                                try
                                {
                                    var line = ftab + fp;
                                    ost.WriteLine(line);
                                }
                                catch (UnauthorizedAccessException ex)
                                {
                                    ost.WriteLine("!!! UnauthorizedAccess file: " + ex.Message);
                                }
                            }
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        ost.WriteLine("!!! UnauthorizedAccess directory: " + ex.Message);
                    }
                }
            }
            catch (DirectoryNotFoundException ex)
            {
                ost.WriteLine("!!! DirectoryNotFound: " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ost.WriteLine("!!! UnauthorizedAccess: " + ex.Message);
            }
            catch (PathTooLongException ex)
            {
                ost.WriteLine("!!! PathTooLong: " + ex.Message);
            }
            catch (Exception ex)
            {
                ost.WriteLine("!!! Exception: " + ex.Message);
            }
        }
    }
}

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
                        (a, v) => a.Concat(EnumerateDirectories(v))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            return dirs;
        }

        public static IEnumerable<DirectoryInfo> EnumerateDirectories1(DirectoryInfo di)
        {
            var dirs = Enumerable.Empty<DirectoryInfo>();
            try
            {
                dirs = di.EnumerateDirectories()
                    .Where(v => (v.Attributes & FileAttributes.System) == 0)
                    .Aggregate<DirectoryInfo, IEnumerable<DirectoryInfo>>(
                        dirs.Append(di),
                        (a, v) => a.Union(EnumerateDirectories1(v))
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
                var m = FileAttributes.Hidden | FileAttributes.System;
                var dcol = Directory.EnumerateDirectories(dp)
                    .Where((v) => ((new DirectoryInfo(v)).Attributes & m) == 0);
                var dcnt = dcol.Count();

                var idx2 = 0;
                var btab = tab + ((cnt == 0) ? "" : ((idx + 1) == cnt) ? "    " : "│  ");
                var dtab = tab + (((idx + 1) == cnt) ? "└─" : "├─");

                dirs = dcol
                    .Aggregate<string, IEnumerable<string>>(
                        dirs
                        .Append(dtab + Path.GetFileName(dp)),
                        (a, v) => a.Concat(EnumerateDirectories2(v, ref idx2, dcnt, btab))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
                dirs = dirs.Append("!!! UnauthorizedAccess");
            }
            idx++;
            return dirs;
        }

        public static IEnumerable<string> EnumerateDirectories3(string dp, ref int idx, int cnt = 0, string tab = "")
        {
            var dirs = Enumerable.Empty<string>();
            try
            {
                var m = FileAttributes.Hidden | FileAttributes.System;
                var dcol = Directory.EnumerateDirectories(dp)
                    .Where((v) => ((new DirectoryInfo(v)).Attributes & m) == 0);
                var dcnt = dcol.Count();

                var fcol = Directory.EnumerateFiles(dp)
                    .Where((v) => ((new FileInfo(v)).Attributes & m) == 0)
                    .Select((v) => Path.GetFileName(v));
                if (fcol.Count() > 0) fcol = fcol.Append("");

                var idx2 = 0;
                var btab = tab + ((cnt == 0) ? "" : ((idx + 1) == cnt) ? "    " : "│  ");
                var dtab = tab + (((idx + 1) == cnt) ? "└─" : "├─");
                var ftab = btab + ((dcnt == 0) ? "    " : "│  ");

                dirs = dcol
                    .Aggregate<string, IEnumerable<string>>(
                        dirs
                        .Append(dtab + Path.GetFileName(dp))
                        .Concat(fcol.Select((v) => ftab + v)),
                        (a, v) => a.Concat(EnumerateDirectories3(v, ref idx2, dcnt, btab))
                        );
            }
            catch (System.UnauthorizedAccessException)
            {
                dirs = dirs.Append("!!! UnauthorizedAccess");
            }
            idx++;
            return dirs;
        }

        bool FileList(string src, string dst, int mode, bool bTree, bool bSize, bool bDate)
        {
            var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (Directory.Exists(dst)) doc = dst;
            var dt = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var op = Path.Combine(doc, "dirlist_" + dt + ".txt");
            if (File.Exists(op))
            {
                File.Delete(op);
            }
            if (!Directory.Exists(src))
            {
                return false;
            }
            SetMode(bSize, bDate);
            long StartTicks = DateTime.Now.Ticks;
            using (StreamWriter ost = new StreamWriter(op, true))
            {
                var line = Path.GetFileName(src);
                ost.WriteLine(line);
                var bFile = (mode != 2);
                var bDir = (mode != 1);
                if (!bTree)
                {
                    WriteFileDirList(ost, src, bFile, bDir);
                }
                else
                {
                    WriteFileDirTree(ost, src, bFile, bDir);
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
                if (bDir && !bFile)
                {
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
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            ost.WriteLine("!!! UnauthorizedAccess directory: " + ex.Message);
                        }
                    }
                }
                else
                {
                    foreach (var dp in EnumerateDirectories3(p, ref idx))
                    {
                        try
                        {
                            if (bSubDir)
                            {
                                var line = dp;
                                ost.WriteLine(line);
                            }
                            bSubDir = bDir;
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            ost.WriteLine("!!! UnauthorizedAccess directory: " + ex.Message);
                        }
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

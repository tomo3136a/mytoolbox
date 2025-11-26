using System;
using System.Collections.Generic;
using System.IO;

namespace AutoCopy
{
    public class Log
    {
        //////////////////////////////////////////////////////////////////////
        /// ログ出力
        //////////////////////////////////////////////////////////////////////

        public static void WriteOneFileUpdate(string f, int i)
        {
            var d = f.Substring(0, i);
            var g = Path.GetDirectoryName(Path.GetDirectoryName(d));
            g = f.Substring(g.Length + 1);
            g = Path.GetFileNameWithoutExtension(g);
            if (g == "") return;
            g = TargetFolder.GetBiasName(g);
            var n = Path.GetFileName(f);
            var x = Notify.CreateMessage(g, n, f, d);
            Notify.ToastOut(x);
        }

        public static void WriteFilesUpdate(string d, List<string> lst)
        {
            var g = Path.GetDirectoryName(d);
            if (g == null) return;
            g = TargetFolder.GetBiasName(Path.GetFileName(g));
            var msg = Notify.CreateFileListMessage(d, lst);
            var f = Notify.OutLnk("files", msg);
            var len = d.Length;
            var n = "";
            foreach (var a in lst)
            {
                n += " " + a.Substring(len);
            }
            if (n.Length > 0) n = n.Substring(1);
            var x = Notify.CreateMessage(g, n, f, d);
            Notify.ToastOut(x);
        }
    }
}

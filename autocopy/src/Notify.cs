using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace AutoCopy
{
    public partial class Notify
    {
        /// <summary>
        /// エントリ作成
        /// </summary>
        /// <param name="k"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static string ToEntry(string k, string p)
        {
            var src = Path.GetFileName(k);
            var dst = Path.GetFileName(p);
            var s = dst;
            if (src != dst) s = src + " -> " + s;

            var a = new Uri(p);
            var msg = "<a href=\"" + a.ToString() + "\">";
            msg += s;
            msg += "</a><br/>\r\n";
            return msg;
        }

        /// <summary>
        /// メッセージ作成
        /// </summary>
        /// <param name="d"></param>
        /// <param name="files"></param>
        /// <returns></returns>
        public static string CreateFileListMessage(string d, List<string> files)
        {
            //メッセージ作成
            var len = d.Length;
            var p = new Uri(d).ToString();
            var msg = "<html><body>\r\n";
            msg += "<a href=\"" + p + "\">&lt;folder&gt;</a><br/>\r\n";
            foreach (var f in files)
            {
                p = new Uri(f).ToString();
                msg += "<a href=\"" + p + "\">";
                var s = f.Substring(len);
                msg += s; //Path.GetFileName(s);
                msg += "</a><br/>\r\n";
            }
            msg += "</body></html>";
            return msg;
        }

        //リンクファイル書き出し
        public static string OutLnk(string name, string msg)
        {
            //メッセージディレクトリ確認
            var path = Config.GetPath(Config.FileType.LOG);
            var d = Path.GetDirectoryName(path);
            if (d == null) return "";

            //メッセージファイル保存
            var dt = DateTime.Now.ToString("yyyyMMddHHmm");
            path = Path.Combine(d, name + "_" + dt + ".htm");
            using (var fo = new StreamWriter(path, false))
            {
                fo.WriteLineAsync(msg);
            }
            return path;
        }

        /// <summary>
        /// トーストメッセージファイル作成
        /// </summary>
        /// <param name="msg1">メッセージ1</param>
        /// <param name="msg2">メッセージ2</param>
        /// <param name="f">ファイルパス</param>
        /// <param name="d"></param>
        public static string CreateMessage(string msg1, string msg2, string f, string d = "")
        {
            var p = Config.GetPath(Config.FileType.DOCUMENT, false);
            p = Path.GetDirectoryName(p);
            if (p == null) return "";
            p = Path.Combine(p, "message.xml");

            var d2 = d;
            if (d2 == "") d2 = Path.GetDirectoryName(f);
            if (d2 == null) return "";
            var url = (new Uri(d2)).ToString();
            var launch = url;

            var action = new Uri(f).ToString();
            XElement actions = new XElement("actions");
            XElement elm = new XElement("toast",
                new XAttribute("activationType", "protocol"),
                new XAttribute("launch", launch),
                new XElement("visual",
                    new XAttribute("branding", "name"),
                    new XElement("binding",
                        new XAttribute("template", "ToastGeneric"),
                        new XElement("text",
                            new XAttribute("placement", "attribution"),
                            "ファイル変更通知"),
                        new XElement("text", msg1),
                        new XElement("text", msg2)
                    )
                )
            );
            if (!IsIgnoreFileName(f))
            {
                launch = action;
                actions = new XElement("actions",
                    new XElement("action",
                        new XAttribute("activationType", "protocol"),
                        new XAttribute("arguments", action),
                        new XAttribute("content", "開く")
                    ),
                    new XElement("action",
                        new XAttribute("activationType", "protocol"),
                        new XAttribute("arguments", url),
                        new XAttribute("content", "フォルダ")
                    )
                );
                elm.Add(actions);
            }
            (new XDocument(elm)).Save(p);
            return p;
        }
        private static bool IsIgnoreFileName(string name)
        {
            var ext = System.IO.Path.GetExtension(name).ToLower();
            foreach (var s in _ign_lst)
            {
                if (ext == s) return true;
            }
            return false;
        }
        private static List<string> _ign_lst = new List<string>{
            ".com", ".exe", ".bat", ".cmd",
            ".vbs", ".vbe", ".js", ".jse", "wsf", "wsh",
            ".pl", ".wpl", ".cpl", ".ps1"
        };

        /// <summary>
        /// トーストメッセージ通知
        /// </summary>
        /// <param name="p">メッセージファイルパス</param>
        public static void ToastOut(string p)
        {
            string s = @"$doc = Get-Content """ + p + @""" -Encoding UTF8;";
            s += @"[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;";
            s += @"[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] > $null;";
            s += @"$xml = New-Object Windows.Data.Xml.Dom.XmlDocument; $xml.LoadXml($doc);";
            s += @"$toast = [Windows.UI.Notifications.ToastNotification]::new($xml);";
            s += @"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Microsoft.Windows.Explorer').Show($toast);";

            var cmd = new Process();
            cmd.StartInfo.FileName = "PowerShell.exe";
            cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            cmd.StartInfo.Arguments = s;
            cmd.Start();
        }
    }
}

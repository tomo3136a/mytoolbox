using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using System.Xml;

namespace AutoCopy
{
    public partial class Monitor
    {
        /////////////////////////////////////////////////////////////////////

        /// モニタ対象の追加
        public string AddMonitor(string path)
        {
            var dir = "";
            var ptn = "*";
            var mode = "";
            var name = "";
            var fullname = path;
            var conf = Config.GetPath(Config.FileType.CONFIG);

            //モニタ対象がファイルの場合
            if (File.Exists(path))
            {
                FileInfo src = new FileInfo(path);
                mode = "f";
                dir = src.DirectoryName;
                fullname = src.FullName;
            }

            //モニタ対象がディレクトリの場合
            if (Directory.Exists(path))
            {
                DirectoryInfo src = new DirectoryInfo(path);
                mode = "d";
                dir = src.FullName;
                ptn = "*";
                fullname = src.FullName;
            }

            if ("" == name) return "";

            //設定ファイルを表示
            if ("*" == name)
            {
                var p = Config.GetPath(Config.FileType.DEFAULT);
                System.Diagnostics.Process.Start(p);
                return "";
            }

            //設定ファイルがある場合、設定ファイルに登録済みならなにもせず終了
            if (File.Exists(conf))
            {
                foreach (var line in File.ReadAllLines(conf))
                {
                    if (line[0] == '#') continue;
                    var ss = line.Split('\t');
                    if (ss.Length < 4) continue;
                    if (string.Compare(dir, ss[2]) != 0) continue;
                    if (string.Compare(ptn, ss[3]) == 0) return fullname;
                }
            }

            //設定ファイルにモニタ対象を追加
            using (var fo = new StreamWriter(conf, true))
            {
                var dt = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                var line = dir + "\t" + ptn + "\t" + dt;
                fo.WriteLineAsync(mode + "\t" + name + "\t" + line);
            }

            return fullname;
        }

        ///ファイル変更モニタ実施
        public void InvokeMonitor()
        {
            //設定ファイルが無ければ終了
            var conf = Config.GetPath(Config.FileType.CONFIG);
            if (!File.Exists(conf)) return;

            //前回のモニタ結果がある場合、timに読み込み
            var tim = LoadData();
            bool update = false;

            var dic = new Dictionary<string, string>();
            var idx = 0;
            var cnt = 0;

            //設定ごとに調査
            foreach (var line in File.ReadAllLines(conf))
            {
                if (line[0] == '#') continue;
                var ss = line.Split('\t');
                if (ss.Length < 3) continue;
                var mode = ss[0];
                var name = ss[1];
                var dir = ss[2];
                var ptn = ss[3];
                var date = "";
                if (ss.Length > 4) date = ss[4];
                var dst = "";
                if (ss.Length > 5) dst = ss[5];
                idx++;

                //ディレクトリ一覧を取得
                var dis = new List<DirectoryInfo>();
                dis.Add(new DirectoryInfo(dir));
                if (mode == "d")
                {
                    for (var i = 0; i < dis.Count; i++)
                    {
                        foreach (var di in dis[i].EnumerateDirectories())
                        {
                            if (IsIgnoreFileName(di.Name)) continue;
                            dis.Add(di);
                            if (dis.Count > 100) break;
                        }
                        if (dis.Count > 100) break;
                    }
                }

                //ディレクトリを調査し更新されたファイルリストを取得
                var filelist = new List<string>();
                var file = "";
                if (tim.ContainsKey(name)) date = tim[name];
                var last = date;
                foreach (var di in dis)
                {
                    foreach (var fi in di.EnumerateFiles(ptn))
                    {
                        if (IsIgnoreFileName(fi.Name)) continue;
                        var dt = fi.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss");
                        if (String.Compare(date, dt) >= 0) continue;
                        if (!dic.ContainsKey(fi.FullName)) filelist.Add(fi.FullName);
                        dic[fi.FullName] = name;
                        if (String.Compare(last, dt) >= 0) continue;
                        last = dt;
                        file = fi.FullName;
                    }
                }

                //更新ファイルがある場合、トースト通知
                if (filelist.Count > 0)
                {
                    tim[name] = last;
                    update = true;

                    //ファイルリスト作成
                    var msg = "";
                    foreach (var s in filelist)
                    {
                        msg += " " + s.Substring(1 + dir.Length);
                    }

                    //リンクリスト作成
                    var s2 = Notify.CreateFileListMessage(dir, filelist);
                    s2 = Notify.OutLnk(name, s2);

                    string v = "";
                    if (cnt++ > 0) v = cnt.ToString();
                    if (dst == "")
                    {
                        var p = Config.GetPath(Config.FileType.DOCUMENT, false, v);
                        if (filelist.Count > 1)
                        {
                            file = Config.GetPath(Config.FileType.LOG);
                            file = Path.GetDirectoryName(file);
                            if (file == null) continue;
                            file = Path.Combine(name, file, name + ".htm");
                            Notify.CreateMessage(name, msg, file, dir);
                        }
                        else
                        {
                            Notify.CreateMessage(name, msg, "", file);
                        }
                        Notify.ToastOut(p);
                    }
                    else
                    {
                        var si = new FileInfo(file);
                        var di = new FileInfo(Path.Combine(dst, si.Name));
                        if (di.Exists)
                        {
                            var ticks = di.LastWriteTime.Ticks;
                        }
                        si.CopyTo(di.FullName);
                    }

                    //ログ作成
                    // string log = "";
                    // foreach (var s in filelist)
                    // {
                    //     log += last + "\t" + dic[s] + "\t" + s + "\r\n";
                    // }
                    // OutLog(name, log);
                }
            }

            //データが変更されたらデータファイルに保存
            if (update)
            {
                SaveData(tim);
            }
        }

        /////////////////////////////////////////////////////////////////////

        //データファイル読み込み
        static Dictionary<string, string> LoadData()
        {
            var tim = new Dictionary<string, string>();
            var p = Config.GetPath(Config.FileType.DATA, true);
            if (File.Exists(p))
            {
                foreach (var line in File.ReadAllLines(p))
                {
                    if (line[0] == '#') continue;
                    var ss = line.Split('\t');
                    if (ss.Length < 2) continue;
                    tim.Add(ss[0], ss[1]);
                }
            }
            return tim;
        }

        //データファイル書き出し
        static void SaveData(Dictionary<string, string> tim)
        {
            var p = Config.GetPath(Config.FileType.DATA, true);
            var append = false;
            foreach (var k in tim.Keys)
            {
                using (var f = new StreamWriter(p, append))
                {
                    f.WriteLineAsync(k + "\t" + tim[k]);
                    append = true;
                }
            }
        }

        //リンクファイル書き出し
        // static void OutLnk(string name, string dir, List<string> filelist)
        // {
        //     var msg = "<html><body>\r\n";
        //     msg += "<a href=\"&lt;folder&gt;\">" + dir + "</a><br/>\r\n";
        //     foreach (var s in filelist)
        //     {
        //         msg += "<a href=\"" + s + "\">";
        //         msg += s.Substring(1 + dir.Length);
        //         msg += "</a><br/>\r\n";
        //     }
        //     msg += "</body></html>";

        //     var dt = DateTime.Now.ToString("yyyyMMddHHmm");
        //     var path = GetMonitorPath(FileType.LOG);
        //     path = Path.GetDirectoryName(path);
        //     if (path == null) return;
        //     path = Path.Combine(path, name + "_" + dt + ".htm");
        //     using (var fo = new StreamWriter(path, false))
        //     {
        //         fo.WriteLineAsync(msg);
        //     }
        // }

        /// <summary>
        /// ログファイル書き出し
        /// </summary>
        /// <param name="name"></param>
        /// <param name="msg"></param>
        // static void OutLog(string name, string msg)
        // {
        //     var path = GetMonitorPath(FileType.LOG);
        //     using (var fo = new StreamWriter(path, true))
        //     {
        //         fo.WriteLineAsync(msg);
        //     }
        // }

        /////////////////////////////////////////////////////////////////////

        // static Dictionary<string, string> _exts = new Dictionary<string, string>();

        /// <summary>
        /// 無効ファイル拡張子のチェック(無効ファイルから読込)
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        // static bool IsIgnoreExt(string name)
        // {
        //     if (_exts.Count == 0)
        //     {
        //         var ign = GetMonitorPath(FileType.IGNORE);
        //         if (!File.Exists(ign)) return false;
        //         foreach (var line in File.ReadAllLines(ign))
        //         {
        //             var s = line.Trim();
        //             if (s == "") continue;
        //             if (s == ".") continue;
        //             if (s == "..") continue;
        //             _exts.Add(s.ToLower(), s);
        //         }
        //     }

        //     var ext = System.IO.Path.GetExtension(name).ToLower();
        //     if (_exts.ContainsKey(ext)) return true;
        //     return false;
        // }

        /// <summary>
        /// 無効ファイル拡張子のチェック(固定値)
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        static bool IsIgnoreFileName(string name)
        {
            var ext = System.IO.Path.GetExtension(name).ToLower();
            foreach (var s in _ign_lst)
            {
                if (ext == s) return true;
            }
            return false;
        }
        static List<string> _ign_lst = new List<string>{
            ".com", ".exe", ".bat", ".cmd",
            ".vbs", ".vbe", ".js", ".jse", "wsf", "wsh",
            ".pl", ".wpl", ".cpl", ".ps1"
        };
    }
}

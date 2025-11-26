using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Configuration;
using System.Text.RegularExpressions;

namespace AutoCopy
{
    public class Config
    {
        static string _name = "";

        //////////////////////////////////////////////////////////////////////
        /// アプリケーション名
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// アプリケーション名取得
        /// </summary>
        /// <returns></returns>
        public static string GetName()
        {
            if (_name == "")
            {
                var s = AppDomain.CurrentDomain.FriendlyName;
                _name = Path.GetFileNameWithoutExtension(s);
            }
            return _name;
        }

        //////////////////////////////////////////////////////////////////////
        /// パス取得
        //////////////////////////////////////////////////////////////////////

        public enum FileType
        {
            CONFIG,
            IGNORE,
            DATA,
            MESSAGE,
            DOCUMENT,
            LOG,
            DEFAULT
        };

        /// <summary>
        /// パス取得、flag=trueの場合、ファイルが存在しなければからファイル作成
        /// </summary>
        /// <param name="ft">ファイルタイプ</param>
        /// <param name="bcre">作成許可</param>
        /// <param name="code">管理コード</param>
        /// <returns></returns>
        public static string GetPath(FileType ft, bool bcre = false, string code = "")
        {
            var d = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            d = System.IO.Path.Combine(d, GetName());

            var ext = "";
            switch (ft)
            {
                case FileType.CONFIG: ext = ".ini"; break;
                case FileType.DATA: ext = ".txt"; break;
                case FileType.IGNORE: ext = "txt"; break;
                case FileType.MESSAGE: ext = ".msg"; break;
                case FileType.DOCUMENT: ext = ".xml"; break;
                case FileType.LOG: ext = ".log"; break;
            }

            var f = GetName();
            switch (ft)
            {
                case FileType.IGNORE: f = "ignore"; break;
            }
            f += ((code == "") ? "" : ("_" + code)) + ext;

            var p = System.IO.Path.Combine(d, f);
            if (File.Exists(p)) return p;

            if (bcre)
            {
                if (!Directory.Exists(d)) Directory.CreateDirectory(d);
                using (var fo = new StreamWriter(p))
                {
                    fo.WriteLineAsync("");
                }
            }
            return p;
        }

        /// <summary>
        /// ルートパス取得
        /// 優先順： {Appのパス}/, {実行ディレクトリ}/
        /// </summary>
        /// <returns></returns>
        // public static string GetRoot()
        // {
        //     if (_root == "")
        //     {
        //         var p = Directory.GetCurrentDirectory();
        //         var a = Assembly.GetEntryAssembly();
        //         if (a != null)
        //         {
        //             var s = Path.GetDirectoryName(a.Location);
        //             if (s != null) p = s;
        //         }
        //         _root = p;
        //     }
        //     return _root;
        // }

        //////////////////////////////////////////////////////////////////////
        /// 環境変数
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 環境変数反映
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ApplyEnvString(string s)
        {
            Regex re = new Regex("(?:%)\\w*(?:%)");
            var ms = re.Matches(s);
            var res = s;
            foreach (var m in ms)
            {
                var k = ("" + m).Replace("%", "");
                var v = Environment.GetEnvironmentVariable(k);
                res = res.Replace("%" + k + "%", v);
            }
            return res;
        }

        /// <summary>
        /// 環境変数名に変更
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ToEnvString(string s)
        {
            var m = Environment.GetEnvironmentVariables();
            var vs = new List<string>();
            var vm = new Dictionary<string, string>();
            foreach (var k in m.Keys)
            {
                var v = "" + m[k];
                if (v.Contains(";")) continue;
                if (!v.Contains("\\"))
                {
                    switch (("" + k).ToUpper())
                    {
                        case "SYSTEMDRIVE": break;
                        case "USERNAME": break;
                        case "USERDOMAIN": break;
                        case "COMPUTERNAME": break;
                        default: continue;
                    }
                }
                v = ApplyEnvString(v);
                if (!vs.Contains(v))
                {
                    vs.Add(v);
                    vm.Add(v, "" + k);
                }
            }
            var vi = vs.OrderByDescending<string, string>((v) => v.ToUpper())
                .OrderByDescending<string, long>((v) => v.Length)
                .ToList();
            var res = ApplyEnvString(s);
            foreach (var v in vi)
            {
                var s2 = "%" + vm[v] + "%";
                res = res.Replace(v, s2);
            }
            return res;
        }
    }
}

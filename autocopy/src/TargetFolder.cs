//////////////////////////////////////////////////////////////////////
/// 
//////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;

namespace AutoCopy
{
    public class TargetFolder
    {
        const string _ptn_def = ".*";
        const string _ign_def = "^[@_.~]|[.](bak|tmp)$";

        private string _path = "";
        private string _ptn = _ptn_def;
        private string _ign = _ign_def;

        private static int _dt_div = 1;

        private Dictionary<long, List<string>> _szmap = new Dictionary<long, List<string>>();
        private Dictionary<long, List<string>> _dtmap = new Dictionary<long, List<string>>();
        private Dictionary<string, List<string>> _nmmap = new Dictionary<string, List<string>>();

        //////////////////////////////////////////////////////////////////////

        public TargetFolder()
        {
        }
        public TargetFolder(string p)
        {
            _path = p;
        }

        /// <summary>
        /// 有効ファイル名パターンを設定
        /// </summary>
        /// <param name="ptn">パターン</param>
        public void SetPattern(string ptn)
        {
            _ptn = ptn;
        }

        /// <summary>
        /// 無効ファイル名パターン設定
        /// </summary>
        /// <param name="ptn">パターン</param>
        public void SetIgnore(string ptn)
        {
            _ign = ptn;
        }

        /// <summary>
        /// 解析結果を初期化
        /// </summary>
        public void Clear()
        {
            _szmap.Clear();
            _dtmap.Clear();
            _nmmap.Clear();
        }

        /// <summary>
        /// コピー対象探索
        /// </summary>
        /// <param name="ptn">対象パターン</param>
        /// <param name="ign">対象外パターン</param>
        /// <returns></returns>
        public bool Search(string ptn = "", string ign = "")
        {
            ptn = (ptn == "") ? _ptn : ptn;
            ign = (ign == "") ? _ign : ign;
            Clear();

            if (!Directory.Exists(_path)) return false;

            var len = DirectoryNameLength(_path);
            foreach (var f in EnumerateFiles(_path, _ptn, ign))
            {
                var k = f.Substring(len);
                var fi = new FileInfo(f);

                var sz = fi.Length;
                AddListMap(_szmap, sz, k);

                var dt = fi.LastWriteTime.Ticks / _dt_div;
                AddListMap(_dtmap, dt, k);

                var kw = GetBiasName(fi.Name);
                AddListMap(_nmmap, kw, k);
            }
            return true;
        }

        static void AddListMap(Dictionary<long, List<string>> m, long k, string v)
        {
            if (m.Keys.Contains(k))
            {
                m[k].Add(v);
                return;
            }
            var c = new List<string>();
            c.Add(v);
            m.Add(k, c);
        }
        static void AddListMap(Dictionary<string, List<string>> dic, string k, string v)
        {
            if (dic == null) return;
            if (dic.Keys.Contains(k))
            {
                dic[k].Add(v);
                return;
            }
            var c = new List<string>();
            c.Add(v);
            dic.Add(k, c);
        }

        /// <summary>
        /// コピー元を指定してコピー実行
        /// </summary>
        /// <param name="p">コピー元パス</param>
        /// <returns></returns>
        public bool PullIn(string p)
        {
            if (!Directory.Exists(p)) return false;
            var len1 = DirectoryNameLength(p);
            var len2 = DirectoryNameLength(_path);

            var lst = new List<string>();
            foreach (var f in EnumerateCopyFiles(p))
            {
                var k = f.Substring(len1);
                var o = FileCopy(f, _path, k);
                if (o == "")
                {
                    UI.AddMessage("copy fail: " + k);
                    continue;
                }
                lst.Add(o);
            }

            if (lst.Count == 0) return true;
            if (lst.Count == 1)
            {
                Log.WriteOneFileUpdate(lst[0], len2);
            }
            else
            {
                var d = lst[0].Substring(0, len2);

                //var msg = Log.CreateFileListkMessage(d, lst);
                //var f = Log.OutLnk("files", msg);
                Log.WriteFilesUpdate(d, lst);
                //var x = Log.CreateMessage(f, d, f, d);
                //Log.ToastOut(x);
            }

            return true;
        }

        //////////////////////////////////////////////////////////////////////
        /// 更新ファイル取得
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// ディレクトリ文字列長取得
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private static int DirectoryNameLength(string p)
        {
            var i = p.Length;
            if (i == 0) return 0;
            return p.Length + ((p[i - 1] == '\\') ? 0 : 1);
        }

        /// <summary>
        /// 更新対象取得
        /// </summary>
        /// <param name="p">パス</param>
        /// <returns></returns>
        public IEnumerable<string> EnumerateUpdateFiles(string p)
        {
            if (Directory.Exists(p))
            {
                var n = DirectoryNameLength(p);
                foreach (var f in EnumerateCopyFiles(p))
                {
                    var s = f.Substring(n);
                    yield return s;
                }
            }
        }

        /// <summary>
        /// コピー対象ファイル取得
        /// </summary>
        /// <param name="p">パス</param>
        /// <returns></returns>
        public IEnumerable<string> EnumerateCopyFiles(string p)
        {
            var n = DirectoryNameLength(p);
            foreach (var f in EnumerateFiles(p, _ptn, _ign))
            {
                var s = f.Substring(n);
                var fi = new FileInfo(f);

                // サイズが新規ならコピー実施
                var sz = fi.Length;
                if (!_szmap.ContainsKey(sz))
                {
                    yield return f;
                    continue;
                }

                // 日付が新規ならコピー実施
                var dt = fi.LastWriteTime.Ticks / _dt_div;
                if (!_dtmap.ContainsKey(dt))
                {
                    yield return f;
                    continue;
                }

                // カテゴリが新規ならコピー実施
                var kw = GetBiasName(Path.GetFileName(s));
                if (!_nmmap.Keys.Contains(kw))
                {
                    yield return f;
                    continue;
                }

                // サイズが同じでカテゴリも同じで日付も同じならコピーしない
                var skip = false;
                foreach (var s1 in _szmap[sz])
                {
                    if (_nmmap[kw].Contains(s1))
                    {
                        if (_dtmap[dt].Contains(s1))
                        {
                            skip = true;
                            break;
                        }
                    }
                }
                if (skip) continue;

                yield return f;
            }
        }

        //////////////////////////////////////////////////////////////////////
        /// ファイル一覧取得
        //////////////////////////////////////////////////////////////////////

        const FileAttributes _fmask = FileAttributes.Hidden | FileAttributes.System;

        /// <summary>
        /// ファイル一覧取得
        /// </summary>
        /// <param name="p"></param>
        /// <param name="ptn"></param>
        /// <param name="ign"></param>
        /// <returns></returns>
        static IEnumerable<string> EnumerateFiles(string p, string ptn = "", string ign = "")
        {
            var re_ptn = new Regex(ptn == "" ? _ptn_def : ptn, RegexOptions.IgnoreCase);
            var re_ign = new Regex(_ign_def + (ign == "" ? "" : ("|" + ign)), RegexOptions.IgnoreCase);

            foreach (var d in EnumerateDirectories(p))
            {
                foreach (var f in Directory.EnumerateFiles(d))
                {
                    if ((File.GetAttributes(f) & _fmask) != 0) continue;
                    var n = Path.GetFileName(f);
                    if (ptn != "") if (!re_ptn.IsMatch(n)) continue;
                    if (re_ign.IsMatch(n)) continue;
                    if (File.Exists(Path.Combine(d, "~$" + n))) continue;
                    yield return f;
                }
            }
        }

        const string _dign = "^[@_.~]";
        private static Regex _re_dign = new Regex(_dign);

        /// <summary>
        /// ディレクトリ一覧取得
        /// </summary>
        /// <param name="p">パス</param>
        /// <returns></returns>
        static IEnumerable<string> EnumerateDirectories(string p)
        {
            yield return p;
            IEnumerable<string> ds = Directory.EnumerateDirectories(p);
            foreach (var d in ds)
            {
                if ((File.GetAttributes(d) & _fmask) != 0) continue;
                if (_re_dign.IsMatch(Path.GetFileName(d))) continue;
                foreach (var r in EnumerateDirectories(d))
                {
                    yield return r;
                }
            }
        }

        //////////////////////////////////////////////////////////////////////
        /// ファイルコピー
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// ファイルコピー
        // 送り先にファイルがなければコピー
        // ディレクトリがなければディレクトリも作成
        // コピー元(src)のほうが古ければ、日付付きリネームしたコピー先(dst)にコピー
        // コピー元(src)のほうが新しければ、コピー先(dst)を日付付きリネーム後コピー
        /// </summary>
        /// <param name="src"></param>
        /// <param name="d"></param>
        /// <param name="f"></param>
        /// <returns></returns>
        static string FileCopy(string src, string d, string f)
        {
            var dst = Path.Combine(d, f);
            try
            {
                if (File.Exists(dst))
                {
                    var t1 = File.GetLastWriteTime(src);
                    var t2 = File.GetLastWriteTime(dst);
                    var tt1 = t1.Ticks / _dt_div;
                    var tt2 = t2.Ticks / _dt_div;

                    if (tt1 < tt2)
                        dst = GetNextPath(dst, t1);
                    else if (tt1 > tt2)
                        File.Move(dst, GetNextPath(dst, t2));
                }
                else
                {
                    var d2 = Path.GetDirectoryName(dst);
                    if (d2 == null) return "";
                    if (!Directory.Exists(d2))
                        Directory.CreateDirectory(d2);
                }

                File.Copy(src, dst);
            }
            catch (IOException) { return ""; }
            return dst;
        }

        //////////////////////////////////////////////////////////////////////
        /// 索引

        private const string _ptn_index = "^(\\d{8}[a-z]*(?=_))|((?=_)\\d{8}[a-z]*(?=[.]))";
        private static Regex _re_index = new Regex(_ptn_index);

        public static bool HasIndex(string f)
        {
            return _re_index.IsMatch(f);
        }
        public static bool HasPreIndex(string f)
        {
            var m = _re_index.Match(f);
            if (!m.Success) return false;
            return m.Groups[1].Value != "";
        }
        public static bool HasPostIndex(string f)
        {
            var m = _re_index.Match(f);
            if (!m.Success) return false;
            return m.Groups[2].Value != "";
        }

        public static string GetIndex(string f)
        {
            return _re_index.Match(f).Value;
        }

        public static string RemoveIndex(string f)
        {
            if (!_re_index.IsMatch(f)) return f;
            return _re_index.Replace(f, "");
        }

        public static string UpdateIndex(string f, string index)
        {
            var s = RemoveIndex(f);
            if (HasPreIndex(f)) return index + "_" + f;
            var ext = Path.GetExtension(s);
            s = s.Substring(0, ext.Length);
            return s + index + ext;
        }

        /// <summary>
        /// 日付付き新規ファイル名取得
        /// </summary>
        /// <param name="p">ファイルパス</param>
        /// <param name="dt">日付</param>
        /// <returns></returns>
        static string GetNextPath(string p, DateTime dt)
        {
            if (!File.Exists(p)) return p;

            var d = Path.GetDirectoryName(p);
            if (d == null) return "";

            var f = Path.GetFileName(p);
            var ext = Path.GetExtension(f);
            f = GetBiasName(f);
            f = Path.GetFileNameWithoutExtension(f);

            f += "_" + dt.ToString("yyyyMMdd");
            var s = "";
            for (var i = 0; ; i++)
            {
                s = f + AlphabetNumber(i) + ext;
                s = Path.Combine(d, s);
                if (!File.Exists(s)) break;
            }
            return s;
        }

        //////////////////////////////////////////////////////////////////////
        /// 基準名称

        private const string _ptn_bias = "^(■|【[^】]*】|\\d{8}[a-z]*_)*|_\\d{8}[a-z]*(?=[.])";
        private static Regex _re_bias = new Regex(_ptn_bias);

        /// <summary>
        /// 基準ファイル名取得
        /// </summary>
        /// <param name="f">ファイル名</param>
        /// <returns></returns>
        public static string GetBiasName(string f)
        {
            return _re_bias.Replace(f, "");
        }

        //////////////////////////////////////////////////////////////////////
        /// ステータス

        private const string _ptn_sts = "^■|^【[^】]*】";
        private static Regex _re_sts = new Regex(_ptn_sts);

        public static bool HasStatus(string f)
        {
            return _re_sts.IsMatch(f);
        }

        public static string GetStatus(string f)
        {
            return _re_bias.Match(f).Value.Replace("【", "").Replace("】", "");
        }

        public static string RemoveStatus(string f)
        {
            if (!_re_sts.IsMatch(f)) return f;
            return _re_bias.Replace(f, "");
        }

        public static string UpdateStatus(string f, string sts)
        {
            var s = RemoveStatus(f);
            if (sts == "") return s;
            if (sts == "■") return sts + s;
            return "【" + sts.Replace("【", "").Replace("】", "") + "】" + s;
        }

        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 10進数から英字番号に変換
        /// </summary>
        /// <param name="n">番号</param>
        /// <returns></returns>
        public static string AlphabetNumber(int n)
        {
            var s = "";
            var a = n - 1;
            while (a >= 0)
            {
                s = (char)(97 + (a % 26)) + s;
                a = a / 26 - 1;
            }
            return s;
        }
    }
}

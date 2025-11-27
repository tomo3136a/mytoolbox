using System;
using System.Collections.Generic;
using System.IO;

namespace AutoCopy
{
    public class Action
    {
        string _path = "";
        string _ign = "";
        string _ptn = "";
        List<string> _src_lst = new List<string>();
        List<string> _dst_lst = new List<string>();

        private TargetFolder _path_info = new TargetFolder();
        private List<TargetFolder> _dst_info = new List<TargetFolder>();

        //////////////////////////////////////////////////////////////////////
        /// 
        //////////////////////////////////////////////////////////////////////

        public Action()
        {
        }

        public Action(string p)
        {
            _path = p;
        }

        public string GetPath()
        {
            return _path;
        }

        public bool AddSource(string p)
        {
            if (_src_lst.Contains(p)) return false;
            _src_lst.Add(p);
            return true;
        }
        public bool RemoveSource(string p)
        {
            if (!_src_lst.Contains(p)) return false;
            _src_lst.Remove(p);
            return true;
        }
        public List<string> GetSources()
        {
            return _src_lst;
        }

        public bool AddDestination(string p)
        {
            if (_dst_lst.Contains(p)) return false;
            _dst_lst.Add(p);
            return true;
        }
        public bool RemoveDestination(string p)
        {
            if (!_dst_lst.Contains(p)) return false;
            _dst_lst.Remove(p);
            return true;
        }
        public List<string> GetDestinations()
        {
            return _dst_lst;
        }

        public void SetIgnore(string ptn)
        {
            _ign = ptn;
        }
        public string GetIgnore()
        {
            return _ign;
        }

        public void SetPattern(string ptn)
        {
            _ptn = ptn;
        }
        public string GetPattern()
        {
            return _ptn;
        }

        /// <summary>
        /// 操作設定保存
        /// </summary>
        /// <returns></returns>
        public bool Save()
        {
            var res = false;
            try
            {
                if (_path == "") return false;
                if (!Directory.Exists(_path))
                {
                    Directory.CreateDirectory(_path);
                }
                var n = "." + Config.GetName().ToLower();
                var p = Path.Combine(_path, n);
                using (var wst = new StreamWriter(p))
                {
                    foreach (var s in _src_lst)
                    {
                        if (s == _path) continue;
                        if (_dst_lst.Contains(s))
                            wst.WriteLine("syn: " + s);
                        else
                            wst.WriteLine("src: " + s);
                    }
                    foreach (var s in _dst_lst)
                    {
                        if (s == _path) continue;
                        if (_src_lst.Contains(s)) continue;
                        wst.WriteLine("dst: " + s);
                    }
                    if (_ign != "") wst.WriteLine("ign: " + _ign);
                    if (_ptn != "") wst.WriteLine("ptn: " + _ptn);
                }
                res = true;
            }
            catch (IOException e)
            {
                UI.AddMessage("設定保存：");
                UI.ShowMessage(e.Message);
            }
            return res;
        }

        /// <summary>
        /// 操作設定読み出し
        /// </summary>
        /// <returns></returns>
        public bool Load()
        {
            try
            {
                var n = "." + Config.GetName().ToLower();
                var p = Path.Combine(_path, n);
                if (!File.Exists(p)) return false;
                foreach (var line in File.ReadAllLines(p))
                {
                    var ss = line.Split(new char[] { ':' }, 2);
                    if (ss.Length < 2) continue;
                    var k = ss[0].Trim().ToLower();
                    var v = ss[1].Trim();
                    switch (k)
                    {
                        case "syn": AddSource(v); AddDestination(v); break;
                        case "src": AddSource(v); break;
                        case "dst": AddDestination(v); break;
                        case "ign": _ign = v; break;
                        case "ptn": _ptn = v; break;
                    }
                }
                return true;
            }
            catch (IOException e)
            {
                UI.AddMessage("操作設定読み出し：");
                UI.ShowMessage(e.Message);
            }
            return false;
        }

        //////////////////////////////////////////////////////////////////////
        /// 検索
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 取り込みアイテム検索
        /// </summary>
        /// <returns></returns>
        public bool FindFrom()
        {
            if (_path == "") return false;
            _path_info = new TargetFolder(_path);
            _path_info.SetIgnore(_ign);
            _path_info.SetPattern(_ptn);
            return _path_info.Search();
        }

        /// <summary>
        /// 吐き出しアイテム検索
        /// </summary>
        /// <returns></returns>
        public bool FindTo()
        {
            _dst_info.Clear();
            foreach (var s in _dst_lst)
            {
                if (s == "") continue;
                if (s == _path) return false;
                var info = new TargetFolder(s);
                info.SetIgnore(_ign);
                info.SetPattern(_ptn);
                info.Search();
                _dst_info.Add(info);
            }
            return true;
        }

        /// <summary>
        /// 更新ファイル一覧取得
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public IEnumerable<string> EnumerateSourceUpdateFiles(string p)
        {
            return _path_info.EnumerateUpdateFiles(p);
        }

        /// <summary>
        /// 更新ファイル一覧取得
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>

        public IEnumerable<string> EnumerateDestinationUpdateFiles(string p)
        {
            var lst = new List<string>();
            foreach (var tgt in _dst_info)
            {
                foreach (var f in tgt.EnumerateUpdateFiles(p))
                {
                    if (lst.Contains(f)) continue;
                    lst.Add(f);
                }
            }
            foreach (var f in lst)
            {
                yield return f;
            }
        }

        //////////////////////////////////////////////////////////////////////
        /// 複製
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 取り込みアイテム複製
        /// </summary>
        /// <returns></returns>
        public bool CopyFrom()
        {
            foreach (var s in _src_lst)
            {
                _path_info.PullIn(s);
            }
            return true;
        }

        /// <summary>
        /// 吐き出しアイテム複製
        /// </summary>
        /// <returns></returns>
        public bool CopyTo()
        {
            var f = true;
            foreach (var tgt in _dst_info)
            {
                f &= tgt.PullIn(_path);
            }
            return f;
        }
    }
}

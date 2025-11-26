using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace AutoCopy
{
    public class Service
    {
        private static Service _serv = new Service();

        string _conf = "";
        List<string> _ent_lst = new List<string>();

        //////////////////////////////////////////////////////////////////////
        /// サービス

        /// <summary>
        /// サービス取得
        /// </summary>
        /// <returns></returns>
        public static Service GetCurrentService()
        {
            return _serv;
        }
        public static Service NewService()
        {
            _serv = new Service();
            return _serv;
        }

        //////////////////////////////////////////////////////////////////////
        /// 設定ファイル

        /// <summary>
        /// 設定ファイルパス設定
        /// </summary>
        /// <param name="p"></param>
        public void SetConfigPath(string p)
        {
            if (_conf != "" && _conf == p) return;
            InitApplication(p);
            _conf = p;
        }

        /// <summary>
        /// 設定ファイルパス取得
        /// </summary>
        /// <returns></returns>
        public string GetConfigPath()
        {
            return _conf;
        }

        //////////////////////////////////////////////////////////////////////
        /// エントリ

        /// <summary>
        /// エントリ追加
        /// </summary>
        /// <param name="p"></param>
        public void AddEntry(string p)
        {
            _ent_lst.Add(p);
        }
        /// <summary>
        /// エントリファイル読込
        /// </summary>
        /// <param name="p"></param>
        public void LoadEntriesFile(string p)
        {
            if (File.Exists(p))
            {
                try
                {
                    using (var ist = new StreamReader(p))
                    {
                        var s = "";
                        while ((s = ist.ReadLine()) != null)
                        {
                            _ent_lst.Add(s);
                        }
                    }
                }
                catch (IOException e)
                {
                    UI.ShowMessage(e.Message);
                }
            }
        }

        //////////////////////////////////////////////////////////////////////
        /// アプリケーション設定
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// 設定初期化
        /// </summary>
        public void InitApplication(string p = "")
        {
            Load(p);
        }

        //////////////////////////////////////////////////////////////////////
        /// 設定ファイル
        //////////////////////////////////////////////////////////////////////

        private Dictionary<string, string> _dic = new Dictionary<string, string>();

        /// <summary>
        /// 設定ファイルパス取得
        /// </summary>
        /// <returns></returns>
        public string GetAppConfFile()
        {
            if (_conf == "") Load();
            return _conf;
        }

        /// <summary>
        /// 設定ファイル読込
        /// </summary>
        /// <param name="p"></param>
        public void Load(string p = "")
        {
            _conf = p == "" ? Config.GetPath(Config.FileType.CONFIG) : p;
            _dic.Clear();
            _ent_lst.Clear();
            if (!File.Exists(_conf)) return;

            try
            {
                foreach (var line in File.ReadAllLines(_conf))
                {
                    var ss = line.Split(new char[] { '=' }, 2);
                    if (ss.Length < 2) continue;
                    var k = ss[0].Trim().ToLower();
                    var v = ss[1].Trim();
                    if (_dic.ContainsKey(k)) continue;
                    _dic[k] = Config.ApplyEnvString(v);
                }
                foreach (var v in _dic.Values)
                {
                    if (!_ent_lst.Contains(v)) _ent_lst.Add(v);
                }
            }
            catch (IOException e)
            {
                UI.AddMessage("設定読み出し：");
                UI.ShowMessage(e.Message);
            }
        }

        /// <summary>
        /// 設定ファイル書き出し
        /// </summary>
        /// <param name="p"></param>
        public void Save(string p = "")
        {
            if (p != "") _conf = p;
            if (_conf == "") _conf = Config.GetPath(Config.FileType.CONFIG, true);
            try
            {
                using (var wst = new StreamWriter(_conf))
                {
                    foreach (var k in _dic.Keys)
                    {
                        wst.WriteLine(k + "=" + _dic[k]);
                    }
                }
            }
            catch (IOException e)
            {
                UI.AddMessage("設定書き出し：");
                UI.ShowMessage(e.Message);
            }
        }

        /// <summary>
        /// 設定クリア
        /// </summary>
        /// <returns></returns>
        public bool ClearAppValue()
        {
            _dic.Clear();
            return true;
        }

        /// <summary>
        /// 設定に値設定
        /// </summary>
        /// <param name="k"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        public bool SetAppValue(string k, string v)
        {
            if (_dic.ContainsKey(k)) _dic[k] = v;
            else _dic.Add(k, v);
            return true;
        }

        /// <summary>
        /// 設定から値取得
        /// </summary>
        /// <param name="k"></param>
        /// <returns></returns>
        public string GetAppValue(string k)
        {
            return _dic.ContainsKey(k) ? _dic[k] : "";
        }

        public IEnumerable<string> GetSourceUpdateEnumerator()
        {
            if (_act.FindFrom())
            {
                foreach (var p in _act.GetSources())
                {
                    foreach (var s in _act.EnumerateSourceUpdateFiles(p))
                    {
                        yield return s;
                    }
                }
            }
        }

        public IEnumerable<string> GetDestinationUpdateEnumerator()
        {
            if (_act.FindTo())
            {
                var p = _act.GetPath();
                foreach (var s in _act.EnumerateDestinationUpdateFiles(p))
                {
                    yield return s;
                }
            }
        }

        public void CopySync()
        {
            CopyTo();
            CopyFrom();
        }

        public void CopyFrom()
        {
            UI.ResetMessage();
            if (!_act.CopyFrom())
            {
            }
            UI.ShowMessage();
        }

        public void CopyTo()
        {
            UI.ResetMessage();
            if (!_act.CopyTo())
            {
            }
            UI.ShowMessage();
        }

        //////////////////////////////////////////////////////////////////////
        /// 

        /// <summary>
        /// 初期化
        /// </summary>
        public void Initialize()
        {
            Load(_conf);
        }

        static Action _act = new Action();

        public bool AddSource(string p)
        {
            return _act.AddSource(p);
        }

        public bool AddDestination(string p)
        {
            return _act.AddDestination(p);
        }

        public bool RemoveSource(string p)
        {
            return _act.RemoveSource(p);
        }

        public bool RemoveDestination(string p)
        {
            return _act.RemoveDestination(p);
        }

        public void SetEntry(string p)
        {
            _act = new Action(p);
        }
        public Action GetAction()
        {
            return _act;
        }

        private bool _service_flag = false;
        private bool _gui_flag = false;

        public void SetServiceMode(bool f = true)
        {
            _service_flag = f;
        }
        public void SetGuiMode(bool f = true)
        {
            _gui_flag = f;
        }

        /// <summary>
        /// サービス実行有無
        /// </summary>
        /// <returns></returns>
        public bool IsService()
        {
            if (_gui_flag) return false;
            if (_service_flag) return true;
            return _ent_lst.Count > 0;
        }

        /// <summary>
        /// 実行
        /// </summary>
        public void Action()
        {
            if (_conf == "")
            {
                _conf = GetAppConfFile();
                LoadEntriesFile(_conf);
            }
            foreach (var p in _ent_lst)
            {
                _act = new Action(p);
                if (_act.Load())
                {
                    if (!_act.FindTo())
                    {
                        UI.ShowMessage("Error Find to: " + p);
                    }
                    if (!_act.CopyTo())
                    {
                        UI.ShowMessage("Error Copy To: " + p);
                    }
                    if (!_act.FindFrom())
                    {
                        UI.ShowMessage("Error Find from: " + p);
                    }
                    if (!_act.CopyFrom())
                    {
                        UI.ShowMessage("Error Copy From: " + p);
                    }
                }
            }
        }
    }
}
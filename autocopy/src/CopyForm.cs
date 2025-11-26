using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace AutoCopy
{
    public partial class CopyForm : Form
    {
        static Service _serv = Service.GetCurrentService();

        public CopyForm()
        {
            InitializeComponent();

            LoadConfig();
            LoadEntry();
        }

        public void UpdateTitle()
        {
            this.Text = Config.GetName() + " - " + _serv.GetAppConfFile();
        }

        //////////////////////////////////////////////////////////////////////
        /// イベントハンドラ
        //////////////////////////////////////////////////////////////////////

        void OnClose(Object sender, EventArgs e)
        {
            this.Close();
        }

        void OnSave(Object sender, EventArgs e)
        {
            SaveConfig();
            SaveEntry();
        }

        void OnChangedPath(Object sender, EventArgs e)
        {
            var cmb = (ComboBox)sender;
            if (cmb.Name == "e")
            {
                _serv.SetEntry(cmb_ent.Text);
                LoadEntry();
            }
        }

        void OnRemovePath(Object sender, EventArgs e)
        {
            var btn = (Button)sender;
            if (btn.Name == "e") RemoveEntry();
            else if (btn.Name == "p") RemovePath();
        }

        void OnAddPath(Object sender, EventArgs e)
        {
            var btn = (Button)sender;
            if (btn.Name == "e") AddEntry();
            else if (btn.Name == "p") AddPath();
        }

        void OnSelectPath(Object sender, EventArgs e)
        {
            var btn = (Button)sender;
            if (btn.Name == "c") SetConfig();
        }

        void OnEditPath(Object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var p = "";
            if (btn.Name == "e")
            {
                p = _serv.GetAppConfFile();
            }
            else if (btn.Name == "p")
            {
                var n = "." + Config.GetName().ToLower();
                p = Path.Combine(cmb_ent.Text, n);
            }
            if (p == "") return;

            var pi = new ProcessStartInfo();
            pi.FileName = "notepad";
            pi.Arguments = p;
            pi.ErrorDialog = true;
            pi.UseShellExecute = true;
            Process.Start(pi);
        }

        void OnFind(Object sender, EventArgs e)
        {
            UpdateSourceTree();
            UpdateDestinationTree();
        }

        void OnCopySync(Object sender, EventArgs e)
        {
            SaveEntry();
            _serv.CopySync();
            UpdateSourceTree();
            UpdateDestinationTree();
        }

        void OnCopyFrom(Object sender, EventArgs e)
        {
            SaveEntry();
            _serv.CopyFrom();
            UpdateSourceTree();
        }

        void OnCopyTo(Object sender, EventArgs e)
        {
            SaveEntry();
            _serv.CopyTo();
            UpdateDestinationTree();
        }

        //////////////////////////////////////////////////////////////////////
        /// パスツリー

        /// <summary>
        /// パスツリー更新
        /// </summary>
        private void UpdateSourceTree()
        {
            tvw_src.Nodes.Clear();
            foreach (var s in _serv.GetSourceUpdateEnumerator())
            {
                tvw_src.Nodes.Add(s);
            }
        }

        /// <summary>
        /// パスツリー更新
        /// </summary>
        private void UpdateDestinationTree()
        {
            tvw_dst.Nodes.Clear();
            foreach (var s in _serv.GetDestinationUpdateEnumerator())
            {
                tvw_dst.Nodes.Add(s);
            }
        }

        //////////////////////////////////////////////////////////////////////
        /// エントリ

        /// <summary>
        /// エントリ削除
        /// </summary>
        void RemoveEntry()
        {
            var n = cmb_ent.Items.IndexOf(cmb_ent.Text);
            if (n >= 0) cmb_ent.Items.RemoveAt(n);
            cmb_ent.Text = "";
            n = (n < cmb_ent.Items.Count) ? n : (cmb_ent.Items.Count - 1);
            if (n >= 0) cmb_ent.SelectedIndex = n;
        }

        /// <summary>
        /// エントリ追加
        /// </summary>
        /// <param name="p"></param>
        void AddEntry(string p)
        {
            if (p == "") return;
            if (p == cmb_ent.Text) return;
            if (cmb_ent.Items.Contains(p)) return;
            cmb_ent.Items.Add(p);
            cmb_ent.Text = p;
        }
        void AddEntry()
        {
            var p = cmb_ent.Text;
            if (UI.SelectFolder(ref p)) AddEntry(p);
        }

        void SetConfig()
        {
            var p = _serv.GetAppConfFile();
            var s = UI.SelectFile(p);
            if (s == "") return;
            _serv.InitApplication(s);
            LoadConfig();
            UpdateTitle();
        }


        //////////////////////////////////////////////////////////////////////
        /// アイテムパス

        /// <summary>
        /// アイテムパス削除
        /// </summary>
        void RemovePath()
        {
            var f = false;
            foreach (DataGridViewRow r in gvw_path.SelectedRows)
            {
                var v = r.Cells[0].Value;
                if (v == null) continue;
                var p = (string)v;
                v = r.Cells[1].Value;
                if (v != null)
                    if ((bool)v)
                        f |= _serv.RemoveSource(p);
                v = r.Cells[2].Value;
                if (v != null)
                    if ((bool)v)
                        f |= _serv.RemoveDestination(p);
            }
            if (f)
            {
                UpdatePath();
                SaveEntry();
            }
        }

        /// <summary>
        /// アイテムパス追加
        /// </summary>
        void AddPath()
        {
            var p = "";
            if (gvw_path.SelectedCells.Count > 0)
            {
                var v = gvw_path.SelectedCells[0].Value;
                if (v != null) p = (string)v;
            }
            if (!UI.SelectFolder(ref p)) return;
            if (p == "") return;
            if (_serv.AddSource(p))
            {
                UpdatePath();
                SaveEntry();
            }
        }

        void UpdatePath()
        {
            gvw_path.Rows.Clear();
            var lst = new List<string>();
            var src = new List<string>();
            var dst = new List<string>();
            var act = _serv.GetAction();
            foreach (var s in act.GetSources())
            {
                if (!lst.Contains(s)) lst.Add(s);
                src.Add(s);
            }
            foreach (var s in act.GetDestinations())
            {
                if (!lst.Contains(s)) lst.Add(s);
                dst.Add(s);
            }
            foreach (var s in lst)
                gvw_path.Rows.Add(s, src.Contains(s), dst.Contains(s));
            txt_ign.Text = act.GetIgnore();
        }

        //////////////////////////////////////////////////////////////////////
        /// エントリ設定ファイル

        /// <summary>
        /// エントリ設定ファイル読込
        /// </summary>
        void LoadEntry()
        {
            var act = _serv.GetAction();

            gvw_path.Rows.Clear();
            txt_ign.Text = "";
            if (act.Load())
            {
                UpdatePath();
                txt_ign.Text = act.GetIgnore();
            }
        }

        /// <summary>
        /// エントリ設定ファイル保存
        /// </summary>
        void SaveEntry()
        {
            var act = _serv.GetAction();

            foreach (DataGridViewRow r in gvw_path.Rows)
            {
                var v = r.Cells[0].Value;
                if (v == null) continue;
                var p = v.ToString();
                if (p == null) continue;
                v = r.Cells[1].Value;
                if (v != null)
                    if ((bool)v)
                        act.AddSource(p);
                    else
                        act.RemoveSource(p);
                v = r.Cells[2].Value;
                if (v != null)
                    if ((bool)v)
                        act.AddDestination(p);
                    else
                        act.RemoveDestination(p);
            }
            act.SetIgnore(txt_ign.Text);
            act.Save();
        }

        //////////////////////////////////////////////////////////////////////
        /// 設定ファイル

        /// <summary>
        /// 設定ファイル読込
        /// </summary>
        void LoadConfig()
        {
            cmb_ent.Items.Clear();
            for (var i = 0; i < 10; i++)
            {
                var si = _serv.GetAppValue("cfg" + i);
                if (si == "") continue;
                var pi = Config.ApplyEnvString(si);
                AddEntry(pi);
            }
            var s = _serv.GetAppValue("cfg");
            if (s == "") return;
            var p = Config.ApplyEnvString(s);
            AddEntry(p);
            cmb_ent.Text = p;
            _serv.SetEntry(p);
        }

        /// <summary>
        /// 設定ファイル保存
        /// </summary>
        void SaveConfig()
        {
            _serv.ClearAppValue();
            for (var i = 0; i < cmb_ent.Items.Count; i++)
            {
                var si = cmb_ent.Items[i];
                if (si == null) continue;
                var pi = Config.ToEnvString((string)si);
                _serv.SetAppValue("cfg" + i, pi);
            }
            var s = cmb_ent.Text;
            if (s == null) return;
            var p = Config.ToEnvString((string)s);
            _serv.SetAppValue("cfg", p);
            _serv.Save();
        }
    }
}

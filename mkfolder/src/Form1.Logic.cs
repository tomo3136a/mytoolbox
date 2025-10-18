using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Reflection.Metadata;

namespace mkfolder
{
    partial class Form1
    {
        void UpdatePackage()
        {
            try
            {
                var da = Directory.EnumerateDirectories(root);
                da = da.Select((v) => Path.GetFileName(v))
                    .Where((v) => !(v.StartsWith(".") || v.StartsWith("_")));
                this.lbx_pkg.BeginUpdate();
                lbx_pkg.Items.Clear();
                foreach (var d in da)
                {
                    lbx_pkg.Items.Add(d);
                }
                if (lbx_pkg.Items.Count > 0)
                {
                    lbx_pkg.SelectedIndex = 0;
                }
                this.lbx_pkg.EndUpdate();
            }
            catch (Exception)
            {

            }
        }

        void UpdateSource(string p, string n)
        {
            var p1 = Path.Combine(p, n);
            lbx_dir.BeginUpdate();
            lbx_dir.Items.Clear();
            try
            {
                var di = new DirectoryInfo(p1);
                var cnt = di.GetDirectories().Count();
                foreach (var di1 in di.EnumerateDirectories())
                {
                    lbx_dir.Items.Add(di1.Name);
                    var vs = EnumerateDirectoriesTree(di1);
                    cnt += vs.Count();
                    if (cnt > 50)
                    {
                        lbx_dir.Items.Clear();
                        lbx_dir.Items.Add("...");
                        return;
                    }
                    foreach (var v in vs)
                    {
                        lbx_dir.Items.Add(v);
                    }
                }
            }
            catch (Exception)
            {

            }
            lbx_dir.EndUpdate();
        }

        bool MakeFolder(string src, string dst)
        {
            if (dst.StartsWith(src)) return false;
            if (src.StartsWith(dst)) return false;
            var di_src = new DirectoryInfo(src);
            var di_dst = new DirectoryInfo(dst);
            var chk = (di_dst.GetDirectories().Count() > 0) ||
                      (di_dst.GetFiles().Count() > 0);
            if (chk)
            {
                var msg = "出力先のフォルダは空ではありません。";
                msg += "\n続けますか。\n" + dst;
                var res = MessageBox.Show(msg, app_name, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (res != DialogResult.Yes) return false;
            }
            MakeSubDirectory(di_src, di_dst);
            return true;
        }

        void MakeSubDirectory(DirectoryInfo src, DirectoryInfo dst)
        {
            foreach (var fi in src.EnumerateFiles())
            {
                var f = Path.Combine(dst.FullName, fi.Name);
                if (!File.Exists(f))
                {
                    fi.CopyTo(f);
                }
            }
            foreach (var ds in src.EnumerateDirectories())
            {
                var dd = dst.CreateSubdirectory(ds.Name);
                MakeSubDirectory(ds, dd);
            }
        }

        private static FileAttributes l_mask = FileAttributes.Hidden | FileAttributes.System;

        public static IEnumerable<string> EnumerateDirectoriesTree(DirectoryInfo di, string tab = "", bool bEnd = true)
        {
            IEnumerable<DirectoryInfo> dia = null;
            var dcnt = 0;
            try
            {
                dia = di.EnumerateDirectories().Where((v) => (v.Attributes & l_mask) == 0);
                dcnt = dia.Count();
            }
            catch (System.UnauthorizedAccessException)
            {
            }
            if (dcnt > 0)
            {
                var idx = 0;
                var cnt = dia.Count();
                foreach (var a in dia)
                {
                    idx++;
                    var btab = tab + ((idx == cnt) ? "└─" : "├─");
                    if (a != null) yield return (btab + a.Name);
                    IEnumerable<string> dib = null;
                    try
                    {
                        var dtab = tab + ((cnt == 0) ? "" : (idx == cnt) ? "    " : "│  ");
                        dib = EnumerateDirectoriesTree(a, dtab, (idx == cnt));
                    }
                    catch (System.UnauthorizedAccessException)
                    {
                    }
                    if (dib != null)
                    {
                        foreach (var b in dib)
                        {
                            yield return b;
                        }
                    }
                }
            }
        }

        void LoadConfig(string name)
        {
            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", app_name + ".config");
            GetConfig("root", ref root);
        }

        void SaveConfig()
        {
            try
            {
                var cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                SetConfig(cfg, "root", ToEnvString(root));
                cfg.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(cfg.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
            }
        }

        void GetConfig(string k, ref string v)
        {
            var s = "";
            try
            {
                s = ConfigurationManager.AppSettings[k];
            }
            catch (ConfigurationErrorsException)
            {
            }
            if (s != null) v = ApplyEnvString(s);
        }

        void SetConfig(Configuration cfg, string k, string v)
        {
            var m = cfg.AppSettings.Settings;
            if (m[k] == null)
            {
                m.Add(k, v);
            }
            else
            {
                m[k].Value = v;
            }
        }

        string ApplyEnvString(string s)
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

        string ToEnvString(string s)
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

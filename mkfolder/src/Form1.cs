using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace mkfolder
{
    public partial class Form1 : Form
    {
        static string root = Directory.GetCurrentDirectory();
        static string app_name = AppDomain.CurrentDomain.FriendlyName;

        public Form1()
        {
            var p = Path.Combine(root, "..\\lib\\" + app_name);
            if (Directory.Exists(p)) root = p;
            LoadConfig(app_name);
            InitializeComponent();
            this.tbx_pkg.Visible = false;
            this.tbx_pkg.Text = root;
            this.lbx_pkg.Select();
        }

        void OnMake(Object sender, EventArgs e)
        {
            var ip = "" + lbx_pkg.Items[this.lbx_pkg.SelectedIndex];
            ip = Path.Combine(root, ip);
            var op = SelectFolder("");
            if (op != "")
            {
                SaveConfig();
                if (MakeFolder(ip, op))
                {
                    this.Close();
                }
            }
        }

        void OnClose(Object sender, EventArgs e)
        {
            this.Close();
        }

        void OnSelectPkg(Object sender, EventArgs e)
        {
            SelectFolder((Button)sender, tbx_pkg);
            this.lbx_pkg.Select();
        }

        void OnTextChangePkg(Object sender, EventArgs e)
        {
            root = this.tbx_pkg.Text;
            if (Directory.Exists(root))
            {
                UpdatePackage();
            }
        }

        void OnChangePkg(Object sender, EventArgs e)
        {
            var p = "" + lbx_pkg.Items[this.lbx_pkg.SelectedIndex];
            UpdateSource(root, p);
        }

        void SelectFolder(Button btn, TextBox tbx)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (!Directory.Exists(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            var res = SelectFolder(root);
            if (res != "")
            {
                tbx.Text = res;
                tbx.BackColor = Color.White;
                tbx.BackColor = SystemColors.Window;
            }
        }

        string SelectFolder(string p, string ini = "")
        {
            var res = ini;
            FolderBrowserDialog dlg = new FolderBrowserDialog();
#if (NET6_0_OR_GREATER || NETSTANDARD2_0_OR_GREATER)
            dlg.InitialDirectory = p;
#else
            dlg.SelectedPath = p;
#endif
            if (DialogResult.OK == dlg.ShowDialog())
            {
                res = dlg.SelectedPath;
            }
            dlg.Dispose();
            return res;
        }

        void SelectFile(Button btn, TextBox tbx)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (!Directory.Exists(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            //#if (NET6_0_OR_GREATER || NETSTANDARD2_0_OR_GREATER)
            dlg.InitialDirectory = root;
            //#else
            //            dlg.SelectedPath = root;
            //#endif
            if (DialogResult.OK == dlg.ShowDialog())
            {
                tbx.Text = dlg.FileName;
                tbx.BackColor = Color.White;
                tbx.BackColor = SystemColors.Window;
            }
            dlg.Dispose();
        }
    }
}

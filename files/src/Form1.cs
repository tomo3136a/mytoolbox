using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace files
{
    public partial class Form1 : Form
    {
        const string AppName = "files";
        private long ticks = 0;

        public Form1()
        {
            InitializeComponent();
            this.txt_src.Text = "C:\\";
            this.rb_3.Checked = true;
            this.cb_1.Checked = false;
            ((Button)this.AcceptButton).Select();
        }

        void OnSelect(Object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            var btn = (Button)sender;
            var txt = (btn.Name == "src") ? txt_src : txt_dst;
            var root = txt.Text;
            if (!Directory.Exists(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
#if (NET6_0_OR_GREATER || NETSTANDARD2_0_OR_GREATER)
            dlg.InitialDirectory = root;
#else
            dlg.SelectedPath = root;
#endif
            if (DialogResult.OK == dlg.ShowDialog())
            {
                txt.Text = dlg.SelectedPath;
                txt.BackColor = Color.White;
                txt.BackColor = SystemColors.Window;
            }
            dlg.Dispose();
        }

        void OnClose(Object sender, EventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                var s = this.txt_src.Text;
                if (!Directory.Exists(s))
                {
                    this.txt_src.Focus();
                    this.txt_src.BackColor = Color.Yellow;
                    MessageBox.Show(this, "フォルダを選択してください。", AppName);
                    this.txt_src.SelectAll();
                    return;
                }

                lbl_sts.Text = "処理中";
                this.Update();
                var mode = this.rb_1.Checked ? 1 : rb_2.Checked ? 2 : 3;
                var bTree = this.cb_1.Checked;
                var bSize = this.cb_2.Checked;
                var bDate = this.cb_3.Checked;
                try
                {
                    var src = this.txt_src.Text;
                    var dst = this.txt_dst.Text;
                    if (FileList(src, dst, mode, bTree, bSize, bDate))
                    {
                        string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                        File.AppendAllLines(AppName + ".log", msg.Split('\n'));
                        lbl_sts.Text = msg;
                        this.Update();
                        MessageBox.Show(this, "Successfull.", AppName);
                    }
                    else
                    {
                        string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                        lbl_sts.Text = msg;
                        this.Update();
                        MessageBox.Show(this, "Denied.", AppName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error.\n" + ex.Message, AppName);
                }
            }
            this.Close();
        }
    }
}

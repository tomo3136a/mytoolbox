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
            var txt = txt_src;
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

        void OnTypeSelect(Object sender, EventArgs e)
        {
            if (cb_1.Checked)
            {
                cb_2.Enabled = false;
                cb_3.Enabled = false;
                cb_4.Enabled = false;
                rb_1.Enabled = false;
            }
            else
            {
                cb_2.Enabled = true;
                cb_3.Enabled = true;
                cb_4.Enabled = true;
                rb_1.Enabled = true;
            }
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

                var mode = this.rb_1.Checked ? 1 : rb_2.Checked ? 2 : 3;
                var bTree = this.cb_1.Checked;
                var bSize = this.cb_2.Checked;
                var bDate = this.cb_3.Checked;
                var bAbs = this.cb_4.Checked;
                try
                {
                    var src = this.txt_src.Text;
                    var dst = "";
                    if (!Directory.Exists(dst))
                    {
                        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    }
                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.InitialDirectory = dst;
                    var dt = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var op = "files_" + dt + ".txt";
                    dlg.FileName = op;
                    dlg.Filter = "テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
                    if (DialogResult.OK != dlg.ShowDialog())
                    {
                        dlg.Dispose();
                        return;
                    }
                    dst = dlg.FileName;
                    dlg.Dispose();

                    lbl_sts.Text = "処理中";
                    this.Update();
                    if (FileList(src, dst, mode, bTree, bSize, bDate, bAbs))
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

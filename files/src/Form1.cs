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
            this.txt_1.Text = "C:\\";
            ((Button)this.AcceptButton).Select();
        }

        void OnSelect(Object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            var root = txt_1.Text;
            if (!Directory.Exists(root))
            {
                root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            dlg.SelectedPath = root;
            if (DialogResult.OK == dlg.ShowDialog())
            {
                this.txt_1.Text = dlg.SelectedPath;
                this.txt_1.BackColor = Color.White;
                this.txt_1.BackColor = SystemColors.Window;
            }
            dlg.Dispose();
        }

        void OnClose(Object sender, EventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                var s = this.txt_1.Text;
                if (!Directory.Exists(s))
                {
                    this.txt_1.Focus();
                    this.txt_1.BackColor = Color.Yellow;
                    MessageBox.Show(this, "フォルダを選択してください。", AppName);
                    this.txt_1.SelectAll();
                    return;
                }

                lbl_2.Text = "処理中";
                this.Update();
                var mode = this.rb_1.Checked ? 1 : rb_2.Checked ? 2 : 3;
                var bTree = this.cb_1.Checked;
                var bSize = this.cb_2.Checked;
                var bDate = this.cb_3.Checked;
                if (FileList(this.txt_1.Text, mode, bTree, bSize, bDate))
                {
                    string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                    File.AppendAllLines(AppName + ".log", msg.Split('\n'));
                    lbl_2.Text = msg;
                    this.Update();
                    MessageBox.Show(this, "Successfull.", AppName);
                }
                else
                {
                    string msg = new DateTime(ticks).ToString("HH:mm:ss.FFFFFFF");
                    lbl_2.Text = msg;
                    this.Update();
                    MessageBox.Show(this, "Denied.", AppName);
                }
            }
            this.Close();
        }
    }
}

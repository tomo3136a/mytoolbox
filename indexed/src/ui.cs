using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

namespace Tmm
{
    partial class Program
    {
        /////////////////////////////////////////////////////////////////////
        // dialog

        /// <summary>
        /// input dialog
        /// </summary>
        public class InputDialog : Form
        {
            Label textLabel = new Label();
            Button accept = new Button();
            Button cancel = new Button();
            Button config = new Button();
            ComboBox textBox = new ComboBox();
            Label srcLabel = new Label();
            Label dstLabel = new Label();
            Label modeLabel = new Label();
            ComboBox comboBox = new ComboBox();
            ListBox listBox = new ListBox();

            public InputDialog(string text, string caption, bool bList = false, string sConfig = null)
            {
                int width = 400;
                int height = 190;
                int hList = 0;
                if (bList)
                {
                    hList = 100;
                    height = height + hList;
                }
                this.Width = width;
                this.Height = height;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowIcon = false;
                this.Text = caption;
                this.MinimumSize = new Size(width, height);
                this.StartPosition = FormStartPosition.CenterScreen;
                int w = this.ClientRectangle.Width;
                int h = this.ClientRectangle.Height;
                this.TopMost = true;

                textLabel.Left = 10;
                textLabel.Top = 10;
                textLabel.Text = text;
                textLabel.AutoSize = true;
                textLabel.MaximumSize = new Size(width - 20, 0);

                modeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                modeLabel.Text = "";
                modeLabel.Left = w - 10 - modeLabel.Width;
                modeLabel.Top = 10;
                modeLabel.AutoSize = true;
                modeLabel.Click += new EventHandler(on_mode);
                modeLabel.Visible = false;

                accept.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                accept.Text = "Ok";
                accept.Width = 100;
                accept.Left = w - 2 * (10 + 100);
                accept.Top = h - 10 - 22;
                accept.AutoSize = true;
                accept.Top = h - 10 - accept.Height;
                accept.DialogResult = DialogResult.OK;
                accept.Click += new EventHandler(on_close);

                cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                cancel.Text = "Cancel";
                cancel.Width = 100;
                cancel.Left = w - 10 - 100;
                cancel.Top = h - 10 - cancel.Height;
                cancel.DialogResult = DialogResult.Cancel;
                cancel.Click += new EventHandler(on_close);

                config.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                config.Text = "Abort";
                config.Width = 100;
                config.Left = 10;
                config.Top = h - 10 - cancel.Height;
                config.DialogResult = DialogResult.Abort;
                config.Click += new EventHandler(on_close);

                textBox.Anchor = AnchorStyles.Top | AnchorStyles.Left
                    | AnchorStyles.Right;
                textBox.Width = w - 10 * 2;
                textBox.Left = 10;
                textBox.AutoSize = true;
                textBox.Top = 32;

                srcLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                srcLabel.Left = 10;
                srcLabel.Top = h - 7 * 10 + 5 - accept.Height;
                srcLabel.Text = "src:";
                srcLabel.AutoSize = true;
                srcLabel.MaximumSize = new Size(width - 20, 0);

                dstLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                dstLabel.Left = 10;
                dstLabel.Top = h - 5 * 10 + 10 - accept.Height;
                dstLabel.Text = "dst:";
                dstLabel.AutoSize = true;
                dstLabel.MaximumSize = new Size(width - 20, 0);

                comboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                comboBox.Left = 10;
                comboBox.Top = h - 10 - accept.Height;
                comboBox.AutoSize = true;
                comboBox.Visible = false;

                if (bList)
                {
                    listBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom
                        | AnchorStyles.Left | AnchorStyles.Right;
                    listBox.Width = w - 10 * 2;
                    listBox.Height = hList;
                    listBox.Left = 10;
                    listBox.Top = 60;
                    listBox.AutoSize = true;
                    listBox.Visible = true;
                    listBox.SelectedIndexChanged += new EventHandler(on_changed);
                    listBox.DoubleClick += new EventHandler(on_ok_close);
                }

                this.Controls.Add(textBox);
                this.Controls.Add(srcLabel);
                this.Controls.Add(dstLabel);
                this.Controls.Add(comboBox);
                if (bList)
                {
                    this.Controls.Add(listBox);
                }
                if (null != sConfig)
                {
                    if (sConfig.Length > 0) config.Text = sConfig;
                    this.Controls.Add(config);
                }
                this.Controls.Add(accept);
                this.Controls.Add(cancel);
                this.Controls.Add(textLabel);
                this.Controls.Add(modeLabel);
                this.AcceptButton = accept;
                this.CancelButton = cancel;
                //this.Focus();
            }

            void on_ok_close(Object sender, EventArgs e)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            void on_close(Object sender, EventArgs e)
            {
                this.Close();
            }

            void on_changed(Object sender, EventArgs e)
            {
                this.textBox.Text = this.listBox.Text;
            }


            /////////////////////////////////////////////////////////////////////

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public string Value
            {
                get
                {
                    return textBox.Text;
                }
                set
                {
                    textBox.Text = value;
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public string SrcName
            {
                set
                {
                    srcLabel.Text = value;
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public string DstName
            {
                set
                {
                    dstLabel.Text = value;
                }
            }

            /////////////////////////////////////////////////////////////////////

            int mode;
            public List<string> ModeList = new List<string>();

            void on_mode(Object sender, EventArgs e)
            {
                if (ModeList.Count > 0)
                {

                    ModeIndex = (mode + 1) % ModeList.Count;
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public int ModeIndex
            {
                set
                {
                    if (value < ModeList.Count)
                    {
                        mode = value;
                        int w = ClientRectangle.Width;
                        modeLabel.Text = ModeList[mode];
                        modeLabel.Left = w - 10 - modeLabel.Width;
                    }
                }
            }

            /////////////////////////////////////////////////////////////////////

            public void AddText(string s)
            {
                textBox.Items.Add(s);
                if (s == "") return;
                if (listBox.Items.Contains(s)) return;
                listBox.Items.Add(s);
            }

            public void AddListItem(string s)
            {
                listBox.Items.Add(s);
            }

            public void UpdateList(string kw, string val)
            {
                foreach (var v in Config.GetValues(kw))
                {
                    AddListItem(v);
                }
                AddText("");
                var flg = false;
                foreach (var v in Config.GetValues(kw + @"\recent"))
                {
                    if (v != val) flg = true;
                }
                if (!flg)
                {
                    if (val != null)
                    {
                        if (val.Length > 0) AddText(val);
                    }
                }
                foreach (var v in Config.GetValues(kw + @"\recent"))
                {
                    AddText(v);
                }
            }

            /////////////////////////////////////////////////////////////////////

            public void AddFormatType(string s)
            {
                if (comboBox.Items.Count == 1)
                {
                    comboBox.Visible = true;
                    comboBox.SelectedIndex = 0;
                }
                comboBox.Items.Add(s);
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public string FormatType
            {
                get
                {
                    return comboBox.Text;
                }
                set
                {
                    comboBox.Text = value;
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public int FormatTypeIndex
            {
                set
                {
                    comboBox.SelectedIndex = value;
                }
            }
        }

        /////////////////////////////////////////////////////////////////////
        ///
        /// <summary>
        /// tag input dialog
        /// </summary>
        /// <param name="tag"></param>
        /// <returns></returns>
        public static string TaggingDialog(string tag, string src)
        {
            string title = "indexed";
            string text = "タグを入力してください。";
            InputDialog dlg = new InputDialog(text, title, true);
            dlg.SrcName = "変更前: " + src;
            dlg.DstName = " ";

            string res = null;
            try
            {
                dlg.UpdateList(@"tag", tag);
                dlg.Value = tag;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    res = dlg.Value;
                }
            }
            catch
            {
                MessageBox.Show("operation error. TaggingDialog",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return res;
        }

        /////////////////////////////////////////////////////////////////////

        /// <summary>
        /// comment input dialog
        /// </summary>
        /// <param name="comment"></param>
        /// <returns></returns>
        public static string CommentDialog(string comment, string src)
        {
            string title = "indexed";
            string text = "コメントを入力してください。";
            InputDialog dlg = new InputDialog(text, title, false);
            dlg.SrcName = "変更前: " + src;
            dlg.DstName = " ";

            string res = null;
            try
            {
                dlg.UpdateList(@"note", comment);
                dlg.Value = comment;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    return dlg.Value;
                }
            }
            catch
            {
                MessageBox.Show("operation error.",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return res;
        }

        /////////////////////////////////////////////////////////////////////

        /// <summary>
        /// file rename input dialog
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string RenameDialog(string s)
        {
            while (File.Exists(s) || Directory.Exists(s))
            {
                var p = Path.GetDirectoryName(s);
                var n = Path.GetFileName(s);
                var e = Path.GetExtension(s);
                n = n.Substring(0, n.Length - e.Length);
                var title = "indexed";
                var msg = "フォルダ、またはファイルが存在します。";
                msg += "別名で保存してください。\r\n";
                msg += n + e + "  -> *" + e;
                var dlg = new InputDialog(msg, title);
                dlg.Value = n; //.Substring(0, n.Length - e.Length);
                DialogResult res = dlg.ShowDialog();
                if (res != DialogResult.OK)
                {
                    return null;
                }
                s = Path.Combine(p, dlg.Value + e);
            }
            return s;
        }

        /// <summary>
        /// file save input dialog
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        // public static string SaveDialog(string s)
        // {
        //     while (File.Exists(s) || Directory.Exists(s))
        //     {
        //         string p = Path.GetDirectoryName(s);
        //         string n = Path.GetFileName(s);
        //         string e = Path.GetExtension(s);
        //         n = n.Substring(0, n.Length - e.Length);
        //         string title = "indexed";
        //         string text = "フォルダ 、またはファイルが存在します。";
        //         text += "別名で保存してください。";
        //         InputDialog dlg = new InputDialog(text, title);
        //         dlg.Value = n;
        //         DialogResult res = dlg.ShowDialog();
        //         if (res != DialogResult.OK)
        //         {
        //             return null;
        //         }
        //         s = Path.Combine(p, dlg.Value + e);
        //     }
        //     return s;
        // }

        /////////////////////////////////////////////////////////////////////

        /// <summary>
        /// new name dialog
        /// </summary>
        /// <param name="name"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static InputDialog NewNameDialog(string name, string dt)
        {
            string text = string.Format("説明を入力してください。(日付：{0})", dt);
            InputDialog dlg = new InputDialog(text, "日付フォルダ");
            dlg.Value = name;
            dlg.AddFormatType("<ディレクトリ>");
            dlg.FormatTypeIndex = 0;
            return dlg;
        }

        /////////////////////////////////////////////////////////////////////

        /// <summary>
        /// add Monitor dialog
        /// </summary>
        /// <param name="name"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string AddMonitorDialog(string name, string src, string ptn, string btn = "")
        {
            string title = "ファイル監視";
            string text = "項目名を入力してください。";
            InputDialog dlg = new InputDialog(text, title, false, btn);
            dlg.SrcName = "フォルダ： " + src;
            dlg.DstName = "パターン： " + ptn;

            string res = null;
            try
            {
                //dlg.UpdateList(@"monitor", src);
                dlg.Value = name;
                switch (dlg.ShowDialog())
                {
                    case DialogResult.OK:
                        res = dlg.Value;
                        break;
                    case DialogResult.Abort:
                        res = "*";
                        break;
                }
            }
            catch
            {
                MessageBox.Show("operation error.",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return res;
        }
    }
}

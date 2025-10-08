using System;
using System.Drawing;
using System.Windows.Forms;

namespace files
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        TextBox txt_src = new TextBox();
        Button sel_src = new Button();
        GroupBox grp_1 = new GroupBox();
        RadioButton rb_1 = new RadioButton();
        RadioButton rb_2 = new RadioButton();
        RadioButton rb_3 = new RadioButton();
        CheckBox cb_1 = new CheckBox();
        CheckBox cb_2 = new CheckBox();
        CheckBox cb_3 = new CheckBox();
        CheckBox cb_4 = new CheckBox();
        Label lbl_sts = new Label();

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 210);
            this.MinimumSize = new Size(400, 310);
            this.Text = AppName;

            int w = this.ClientSize.Width;
            int h = this.ClientSize.Height;
            var sz = new Size(w, h);
            var m = 10;

            Label lbl_src = new Label();
            lbl_src.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_src.Text = "フォルダ：";
            lbl_src.Location = new Point(m, m);
            this.Controls.Add(lbl_src);

            txt_src.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            txt_src.Size = new Size(sz.Width - 3 * m - 75, 23);
            txt_src.Location = new Point(m, lbl_src.Location.Y + lbl_src.Height + m * 0);
            this.Controls.Add(txt_src);

            sel_src.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            sel_src.Name = "src";
            sel_src.Text = "選択";
            sel_src.Size = new Size(75, 23);
            sel_src.Location = new Point(sz.Width - 75 - m, lbl_src.Location.Y + lbl_src.Height + m * 0);
            sel_src.Click += new EventHandler(OnSelect);
            this.Controls.Add(sel_src);

            cb_1.Text = "ツリー形式";
            cb_1.Location = new Point(m, txt_src.Location.Y + txt_src.Height + m * 2);
            cb_1.CheckedChanged += new EventHandler(OnTypeSelect);
            this.Controls.Add(cb_1);

            cb_2.Text = "サイズ";
            cb_2.Location = new Point(m, cb_1.Location.Y + cb_1.Height + m * 0);
            this.Controls.Add(cb_2);

            cb_3.Text = "日付";
            cb_3.Location = new Point(m, cb_2.Location.Y + cb_2.Height + m * 0);
            this.Controls.Add(cb_3);

            cb_4.Text = "絶対パス";
            cb_4.Location = new Point(m, cb_3.Location.Y + cb_3.Height + m * 0);
            this.Controls.Add(cb_4);

            grp_1.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left);
            grp_1.Size = new Size(150, 90);
            grp_1.Location = new Point(2 * m + 150, txt_src.Location.Y + txt_src.Height + m);
            this.Controls.Add(grp_1);

            rb_1.Size = new Size(150, 22);
            rb_1.Text = "ファイル";
            rb_1.Location = new Point(m, m);
            rb_1.Checked = true;
            grp_1.Controls.Add(rb_1);

            rb_2.Size = new Size(170, 22);
            rb_2.Text = "フォルダ";
            rb_2.Location = new Point(m, m + (m + 17));
            grp_1.Controls.Add(rb_2);

            rb_3.Size = new Size(170, 22);
            rb_3.Text = "ファイル+フォルダ";
            rb_3.Location = new Point(m, m + (m + 17) * 2);
            grp_1.Controls.Add(rb_3);

            lbl_sts.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left);
            lbl_sts.Text = "";
            lbl_sts.Size = new Size(200, 23);
            lbl_sts.Location = new Point(m, sz.Height - 4 * m);
            this.Controls.Add(lbl_sts);

            Button ok = new Button();
            ok.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
            ok.DialogResult = DialogResult.OK;
            ok.Name = "ok";
            ok.Text = "&OK";
            ok.Size = new Size(75, 23);
            ok.Location = new Point(sz.Width - 80 - 75 - m, sz.Height - 4 * m);
            ok.Click += new EventHandler(OnClose);
            this.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Name = "cancel";
            cancel.Text = "&Cancel";
            cancel.Size = new Size(75, 23);
            cancel.Location = new Point(sz.Width - 75 - m, sz.Height - 4 * m);
            cancel.Click += new EventHandler(OnClose);
            this.Controls.Add(cancel);

            this.AcceptButton = ok;
            this.CancelButton = cancel;
        }

        #endregion
    }
}


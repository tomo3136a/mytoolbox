using System;
using System.Drawing;
using System.Windows.Forms;

namespace mkfolder
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

        TextBox tbx_pkg = new TextBox();
        Button btn_pkg = new Button();
        ListBox lbx_pkg = new ListBox();
        ListBox lbx_dir = new ListBox();

        Label lbl_sts = new Label();

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 450);
            this.Text = app_name;

            var w = this.ClientSize.Width;
            var h = this.ClientSize.Height;
            var sz = new Size(w, h);
            var m = 10;

            Label lbl_pkg = new Label();
            lbl_pkg.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_pkg.Text = "対象：";
            lbl_pkg.Location = new Point(m, m + 3);
            lbl_pkg.Size = new Size(60, 23);
            this.Controls.Add(lbl_pkg);

            tbx_pkg.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            tbx_pkg.Location = new Point(60 + m, m);
            tbx_pkg.Size = new Size(sz.Width - tbx_pkg.Location.X - 75 - m * 2, 23);
            tbx_pkg.TextChanged += new EventHandler(OnTextChangePkg);
            this.Controls.Add(tbx_pkg);

            btn_pkg.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_pkg.Name = "pkg";
            btn_pkg.Text = "選択";
            btn_pkg.Size = new Size(75, 23);
            btn_pkg.Location = new Point(sz.Width - 75 - m, m);
            btn_pkg.Click += new EventHandler(OnSelectPkg);
            this.Controls.Add(btn_pkg);

            this.lbx_pkg.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom);
            this.lbx_pkg.Location = new Point(m, lbl_pkg.Location.Y + lbl_pkg.Height + m * 1);
            this.lbx_pkg.Size = new Size(20 * m, sz.Height - lbx_pkg.Location.Y - 23 - m * 4);
            this.lbx_pkg.SelectedValueChanged += new EventHandler(OnChangePkg);
            this.Controls.Add(this.lbx_pkg);

            this.lbx_dir.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom);
            this.lbx_dir.Location = new Point(lbx_pkg.Location.X + lbx_pkg.Width + m, lbx_pkg.Location.Y);
            this.lbx_dir.Size = new Size(sz.Width - lbx_dir.Location.X - m, sz.Height - lbx_pkg.Location.Y - 23 - m * 4);
            this.Controls.Add(this.lbx_dir);

            lbl_sts.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left);
            lbl_sts.Text = "";
            lbl_sts.Size = new Size(200, 23);
            lbl_sts.Location = new Point(m, sz.Height - 4 * m);
            this.Controls.Add(lbl_sts);

            Button ok = new Button();
            ok.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
            ok.DialogResult = DialogResult.OK;
            ok.Text = "作成(&M)";
            ok.Size = new Size(75, 23);
            ok.Location = new Point(sz.Width - 80 - 75 - m, sz.Height - 4 * m);
            ok.Click += new EventHandler(OnMake);
            this.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Text = "閉じる(&C)";
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

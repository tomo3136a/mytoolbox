using System;
using System.Drawing;
//using System.Security.Cryptography;
using System.Windows.Forms;

namespace AutoCopy
{
    partial class CopyForm
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

        ComboBox cmb_ent = new ComboBox();
        Button btn_add_ent = new Button();
        Button btn_del_ent = new Button();
        Button btn_edt_ent = new Button();

        DataGridView gvw_path = new DataGridView();
        Button btn_add_path = new Button();
        Button btn_del_path = new Button();
        Button btn_edt_path = new Button();

        TextBox txt_ign = new TextBox();
        TextBox txt_ptn = new TextBox();

        TreeView tvw_src = new TreeView();
        TreeView tvw_dst = new TreeView();

        Button btn_sync = new Button();
        Button btn_copyfrom = new Button();
        Button btn_copyto = new Button();

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            int w = 600;
            int h = 460;
            int m = 12;
            int bw = 72;
            int bh = 24;

            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(w, h);
            this.MinimumSize = new System.Drawing.Size(w, 3 * h / 4);
            UpdateTitle();

            var x0 = m;
            var y0 = m;
            var dx = 64;
            var dw = 0;

            var x = x0;
            var y = y0;
            Label lbl_ent = new Label();
            lbl_ent.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_ent.Text = "フォルダ：";
            lbl_ent.Location = new Point(x, y + 6);
            lbl_ent.AutoSize = true;
            this.Controls.Add(lbl_ent);

            x += dx;
            dw = w - x - m - (bw + m) * 3;
            cmb_ent.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            cmb_ent.Name = "e";
            cmb_ent.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_ent.Location = new Point(x, y + 2);
            cmb_ent.Size = new Size(dw, bh);
            cmb_ent.SelectedIndexChanged += new EventHandler(OnChangedPath);
            this.Controls.Add(cmb_ent);

            x += cmb_ent.Width + m;
            btn_add_ent.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_add_ent.Name = "e";
            btn_add_ent.Text = "追加";
            btn_add_ent.Size = new Size(bw, bh);
            btn_add_ent.Location = new Point(x, y);
            btn_add_ent.Click += new EventHandler(OnAddPath);
            this.Controls.Add(btn_add_ent);

            x += btn_add_ent.Width + m;
            btn_del_ent.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_del_ent.Name = "e";
            btn_del_ent.Text = "削除";
            btn_del_ent.Size = new Size(bw, bh);
            btn_del_ent.Location = new Point(x, y);
            btn_del_ent.Click += new EventHandler(OnRemovePath);
            this.Controls.Add(btn_del_ent);

            x += btn_edt_ent.Width + m;
            btn_edt_ent.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_edt_ent.Name = "e";
            btn_edt_ent.Text = "編集";
            btn_edt_ent.Size = new Size(bw, bh);
            btn_edt_ent.Location = new Point(x, y);
            btn_edt_ent.Click += new EventHandler(OnEditPath);
            this.Controls.Add(btn_edt_ent);
            y0 += cmb_ent.Height + m / 2;

            x = x0;
            y = y0;
            Label lbl_path = new Label();
            lbl_path.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_path.Text = "リスト：";
            lbl_path.Location = new Point(x, y + 6);
            lbl_path.AutoSize = true;
            this.Controls.Add(lbl_path);

            x += dx;
            dw = w - x - m - (bw + m) * 1;
            gvw_path.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            gvw_path.Name = "p";
            gvw_path.RowHeadersVisible = false;
            gvw_path.AllowUserToAddRows = false;
            gvw_path.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gvw_path.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gvw_path.Location = new Point(x, y + 2);
            gvw_path.Size = new Size(dw, gvw_path.Height);
            InitializeDataGridView();
            //gvw_path.SelectedIndexChanged += new EventHandler(OnChangedPath);
            this.Controls.Add(gvw_path);

            x += gvw_path.Width + m;
            btn_add_path.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_add_path.Name = "p";
            btn_add_path.Text = "追加";
            btn_add_path.Size = new Size(bw, bh);
            btn_add_path.Location = new Point(x, y);
            btn_add_path.Click += new EventHandler(OnAddPath);
            this.Controls.Add(btn_add_path);

            y += btn_add_path.Height + m / 2;
            btn_del_path.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_del_path.Name = "p";
            btn_del_path.Text = "削除";
            btn_del_path.Size = new Size(bw, bh);
            btn_del_path.Location = new Point(x, y);
            btn_del_path.Click += new EventHandler(OnRemovePath);
            this.Controls.Add(btn_del_path);

            y += btn_del_path.Height + m / 2;
            btn_edt_path.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Right);
            btn_edt_path.Name = "p";
            btn_edt_path.Text = "編集";
            btn_edt_path.Size = new Size(bw, bh);
            btn_edt_path.Location = new Point(x, y);
            btn_edt_path.Click += new EventHandler(OnEditPath);
            this.Controls.Add(btn_edt_path);
            y0 += gvw_path.Height + m / 2;

            x = x0;
            y = y0;
            Label lbl_ign = new Label();
            lbl_ign.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_ign.Text = "例外：";
            lbl_ign.Location = new Point(x, y + 6);
            lbl_ign.AutoSize = true;
            this.Controls.Add(lbl_ign);

            x += dx;
            txt_ign.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            txt_ign.Location = new Point(x, y + 2);
            txt_ign.Size = new Size(dw, bh);
            this.Controls.Add(txt_ign);
            y0 += txt_ign.Height + m;

            x = x0;
            y = y0;
            Label lbl_ptn = new Label();
            lbl_ptn.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left);
            lbl_ptn.Text = "パターン：";
            lbl_ptn.Location = new Point(x, y + 6);
            lbl_ptn.AutoSize = true;
            this.Controls.Add(lbl_ptn);

            x += dx;
            txt_ptn.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            txt_ptn.Location = new Point(x, y + 2);
            txt_ptn.Size = new Size(dw, bh);
            this.Controls.Add(txt_ptn);
            y0 += txt_ptn.Height + m;

            x = x0 + dx;
            y = y0;
            SplitContainer sc = new SplitContainer();
            sc.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
            sc.Location = new Point(x, y);
            sc.Orientation = Orientation.Vertical;
            sc.Size = new Size(dw, h - bh - m * 2 - y);
            this.Controls.Add(sc);

            dw = dw / 2 - 2;
            tvw_src.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
            tvw_src.Size = new Size(dw, h - bh - m * 2 - y);
            sc.SplitterDistance = dw;
            sc.Panel1.Controls.Add(tvw_src);

            tvw_dst.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
            tvw_dst.Size = new Size(dw, h - bh - m * 2 - y);
            sc.Panel2.Controls.Add(tvw_dst);

            x = w - (m + bw) * 7;
            y = h - m - bh;
            Button btn_sel_cfg = new Button();
            btn_sel_cfg.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_sel_cfg.Name = "c";
            btn_sel_cfg.Text = "選択";
            btn_sel_cfg.Size = new Size(bw, bh);
            btn_sel_cfg.Location = new Point(x, y);
            btn_sel_cfg.Click += new EventHandler(OnSelectPath);
            this.Controls.Add(btn_sel_cfg);

            x += m + bw;
            Button btn_save = new Button();
            btn_save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_save.Text = "保存";
            btn_save.AutoSize = true;
            btn_save.Width = bw;
            btn_save.Location = new Point(x, y);
            btn_save.DialogResult = DialogResult.OK;
            btn_save.Click += new EventHandler(OnSave);
            this.Controls.Add(btn_save);

            x += m + bw;
            Button btn_find = new Button();
            btn_find.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_find.Text = "検索";
            btn_find.AutoSize = true;
            btn_find.Width = bw;
            btn_find.Location = new Point(x, y);
            btn_find.DialogResult = DialogResult.OK;
            btn_find.Click += new EventHandler(OnFind);
            this.Controls.Add(btn_find);

            x += m + bw;
            btn_sync.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_sync.Name = "sync";
            btn_sync.Text = "同期実行";
            btn_sync.AutoSize = true;
            btn_sync.Width = bw;
            btn_sync.Location = new Point(x, y);
            btn_sync.DialogResult = DialogResult.OK;
            btn_sync.Click += new EventHandler(OnCopySync);
            this.Controls.Add(btn_sync);

            x += m + bw;
            btn_copyfrom.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_copyfrom.Name = "copyfrom";
            btn_copyfrom.Text = "逆コピー実行";
            btn_copyfrom.AutoSize = true;
            btn_copyfrom.Width = bw;
            btn_copyfrom.Location = new Point(x, y);
            btn_copyfrom.DialogResult = DialogResult.OK;
            btn_copyfrom.Click += new EventHandler(OnCopyFrom);
            this.Controls.Add(btn_copyfrom);

            x += m + bw;
            btn_copyto.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_copyto.Name = "copyto";
            btn_copyto.Text = "コピー実行";
            btn_copyto.AutoSize = true;
            btn_copyto.Width = bw;
            btn_copyto.Location = new Point(x, y);
            btn_copyto.DialogResult = DialogResult.Cancel;
            btn_copyto.Click += new EventHandler(OnCopyTo);
            this.Controls.Add(btn_copyto);

            x += m + bw;
            Button btn_close = new Button();
            btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btn_close.Text = "閉じる";
            btn_close.AutoSize = true;
            btn_close.Width = bw;
            btn_close.Location = new Point(x, y);
            btn_close.DialogResult = DialogResult.Cancel;
            btn_close.Click += new EventHandler(OnClose);
            this.Controls.Add(btn_close);

            this.AcceptButton = btn_sync;
            this.CancelButton = btn_close;
        }

        #endregion

        private void InitializeDataGridView()
        {
            var c = new DataGridViewColumn();
            c.HeaderText = "path";
            c.Width = 200;
            c.CellTemplate = new DataGridViewTextBoxCell();
            c.ReadOnly = true;
            gvw_path.Columns.Add(c);

            c = new DataGridViewCheckBoxColumn();
            c.HeaderText = "from";
            c.Width = 40;
            gvw_path.Columns.Add(c);

            c = new DataGridViewCheckBoxColumn();
            c.HeaderText = "to";
            c.Width = 40;
            gvw_path.Columns.Add(c);
        }
    }
}

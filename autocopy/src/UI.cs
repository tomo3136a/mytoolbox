using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace AutoCopy
{
    public class UI
    {
        //////////////////////////////////////////////////////////////////////
        /// ファイルアイテム選択
        //////////////////////////////////////////////////////////////////////

        /// <summary>
        /// フォルダ選択
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static bool SelectFolder(ref string p)
        {
            bool b = false;
            if (!Directory.Exists(p))
            {
                var ef = Environment.SpecialFolder.MyDocuments;
                p = Environment.GetFolderPath(ef);
            }
            var res = SelectFolder(p);
            if (res != "")
            {
                p = res;
                b = true;
            }
            return b;
        }

        private static string SelectFolder(string p, string ini = "")
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

        /// <summary>
        /// ファイル選択
        /// </summary>
        /// <param name="btn"></param>
        /// <param name="txt"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static bool SelectFile(Button btn, TextBox txt, string p)
        {
            bool b = false;
            if (!File.Exists(p))
            {
                var ef = Environment.SpecialFolder.MyDocuments;
                p = Environment.GetFolderPath(ef);
            }
            var res = SelectFile(p);
            if (res != "")
            {
                txt.Text = res;
                txt.BackColor = Color.White;
                txt.BackColor = SystemColors.Window;
                b = true;
            }
            return b;
        }

        public static string SelectFile(string p, string ini = "")
        {
            var res = ini;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.InitialDirectory = Path.GetDirectoryName(p);
            dlg.FileName = Path.GetFileName(p);
            dlg.CheckFileExists = false;
            if (DialogResult.OK == dlg.ShowDialog())
            {
                res = dlg.FileName;
            }
            dlg.Dispose();
            return res;
        }

        //////////////////////////////////////////////////////////////////////
        /// メッセージ表示
        //////////////////////////////////////////////////////////////////////

        private static string _msg = "";

        /// <summary>
        /// メッセージクリア
        /// </summary>
        public static void ResetMessage()
        {
            _msg = "";
        }

        /// <summary>
        /// メッセージ追加
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="noeol"></param>
        public static void AddMessage(string msg, bool noeol = false)
        {
            _msg += msg;
            if (!noeol) _msg += "\n";
        }

        /// <summary>
        /// メッセージ表示
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="yesno"></param>
        /// <returns></returns>
        public static bool ShowMessage(string msg = "", bool yesno = false)
        {
            if (msg != "") AddMessage(msg, true);
            if (_msg.Length == 0) return false;
            var btn = MessageBoxButtons.OK;
            if (yesno) btn = MessageBoxButtons.YesNo;
            var res = MessageBox.Show(_msg, Config.GetName(), btn);
            _msg = "";
            if (btn == MessageBoxButtons.OK) return true;
            return (res == DialogResult.Yes);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using Tmm;
using System.Runtime.CompilerServices;

internal class Program
{
    private const string _version = "1.0";

    private static bool _verbose = false;
    private static bool _cui = false;
    private bool _append = false;
    private bool UpdateMode { get; set; }
    private bool RelativeFlag = false;
    private bool QuotationFlag = false;
    private bool _g_relative = false;

    string CurrentPath { get; set; }
    string InputPath { get; set; }
    string OutputPath { get; set; }
    string ScriptPath { get; set; }
    string Title { get; set; }
    string Message { get; set; }
    string Command { get; set; }

    string _g_cpath = "";
    string _g_title = "";

    List<string> Items { get; set; }
    List<string> OutLines { get; set; }
    Dictionary<string, string> BaseValue { get; set; }

    private string _name = "";
    private string _value = "";

    Program()
    {
        UpdateMode = false;

        CurrentPath = Environment.CurrentDirectory;
        InputPath = "";
        OutputPath = "";
        ScriptPath = "";

        Title = "";
        Message = "";
        Command = "";

        Items = new List<string>();
        OutLines = new List<string>();
        BaseValue = new Dictionary<string, string>();

        _g_cpath = CurrentPath;
    }

    /// <summary>
    /// global value setting
    /// </summary>
    private void SetGlobal()
    {
        if (Title == "")
        {
            try
            {
                if (InputPath.Length > 0)
                    Title = Path.GetFileNameWithoutExtension(InputPath);
            }
            catch (IOException) { }
            if (Title == "") Title = AppName();
        }
        _g_relative = RelativeFlag;
        _g_cpath = CurrentPath;
        _g_title = Title;
    }

    /// <summary>
    /// reset value form global
    /// </summary>
    private void Reset()
    {
        RelativeFlag = _g_relative;
        CurrentPath = _g_cpath;
        Title = _g_title;
        Command = "";
        Items.Clear();
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// main function
    /// </summary>
    /// <param name="args"></param>
    [STAThread]
    private static void Main(string[] args)
    {
        // system startup initialize
        StartUp();

        // create application
        var app = new Program();

        // command line
        var line = Environment.CommandLine;
        if (app.ParseCommandLine(line) != 0)
        {
            var s1 = "cmd line:" + line;
            system_println("引数が正しくありません。\n" + s1, 3);
            Environment.ExitCode = -1;
            return;
        }
        app.SetGlobal();

        // output file
        var dst = app.OutputPath;
        app.Load(dst);

        var src = app.InputPath;
        if (src == "")
        {
            if (app.Run() < 0)
            {
                Environment.ExitCode = -1;
                return;
            }
        }
        else
        {
            // test update destination file
            if (!app.UpdateMode && File.Exists(dst))
            {
                var src_dt = File.GetLastWriteTime(src);
                var dst_dt = File.GetLastWriteTime(dst);
                if (src_dt < dst_dt) return;
            }

            //run script
            if (!app.RunScript(src))
            {
                var msg = "load script: " + Path.GetFileName(src);
                verbose_println(msg);
                Environment.ExitCode = -1;
                return;
            }
        }

        // save output line data
        if (app.OutLines.Count == 0) return;
        app.Save(dst);
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// startup function
    /// </summary>
    private static void StartUp()
    {
#if (NET6_0_OR_GREATER || NETSTANDARD2_0_OR_GREATER)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
    }

    /// <summary>
    /// Commandline parser
    /// </summary>
    /// <param name="line">commandline</param>
    /// <returns>error code</returns>
    private int ParseCommandLine(string line)
    {
        var args = GetArgs(line);

        var opt_flg = false;
        var opt = "";
        var res = 0;
        foreach (var arg in args)
        {
            var s = arg.Trim();
            if (s.Length == 0) continue;
            if (!opt_flg)
            {
                var re = new Regex(@"^(--?|/)(\w+)(?:(-|\+)?|=(.*))?$");
                var m = re.Match(s);
                if (!m.Success)
                {
                    Items.Add(s);
                    continue;
                }

                opt = m.Groups[2].Value;
                s = m.Groups[4].Value;
                if (m.Groups[1].Value.Length > 1)
                {
                    switch (opt.ToLower())
                    {
                        case "version": return Cmd_Version();
                        case "help": return Cmd_Help();
                        case "verbose": opt = "v"; break;
                        case "console": opt = "C"; break;
                    }
                }
                var b = false;
                switch (opt)
                {
                    case "h": return Cmd_Help();
                    case "v": b = !_verbose; break;
                    case "C": b = !_cui; break;
                    case "u": b = !UpdateMode; break;
                    case "a": b = !_append; break;
                    case "r": b = !RelativeFlag; break;
                    case "q": b = !QuotationFlag; break;
                    case "d": Command = opt; continue; //delete
                    case "p": Command = opt; continue; //prompt
                    case "y": Command = opt; continue; //yesno
                    case "f": Command = opt; continue; //file select
                    case "g": Command = opt; continue; //folder select
                    case "l": Command = opt; continue; //list select
                    case "x": Command = opt; continue;
                }
                switch (m.Groups[3].Value)
                {
                    case "+": b = true; break;
                    case "-": b = false; break;
                }
                switch (opt)
                {
                    case "v": _verbose = b; continue;
                    case "C": _cui = b; continue;
                    case "u": UpdateMode = b; continue;
                    case "a": _append = b; continue;
                    case "r": RelativeFlag = b; continue;
                    case "q": QuotationFlag = b; continue;
                }
                opt_flg = true;
            }
            if (!opt_flg || s.Length == 0) continue;
            s = RemoveQuotation(s);
            s = RemoveEscape(s);
            switch (opt)
            {
                case "i": InputPath = s; break;
                case "o": OutputPath = s; break;
                case "c": CurrentPath = s; break;
                case "s": ScriptPath = s; break;
                case "t": Title = s; break;
                case "m": Message += ((Message == "") ? "" : "\n") + s; break;
                default: res = -1; break;
            }
            opt_flg = false;
        }
        return res;
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// get application name
    /// </summary>
    /// <returns>application name</returns>
    private static string AppName()
    {
        var s = Assembly.GetExecutingAssembly().Location;
        return Path.GetFileNameWithoutExtension(s);
    }

    /// <summary>
    /// system message print line
    /// </summary>
    /// <param name="s">text</param>
    private static void system_println(string msg, long lv = 0, bool noprompt = false)
    {
        if (_cui)
        {
            foreach (var line in msg.Split('\n'))
            {
                var s = line;
                if (!noprompt) s = AppName() + "> " + s;
                switch (lv)
                {
                    case 1: Console.Out.WriteLine(s); break;
                    case 2: Console.Error.WriteLine(s); break;
                    case 3: Console.Error.WriteLine(s); break;
                    default: Console.WriteLine(s); break;
                }
            }
            return;
        }
        var icon = MessageBoxIcon.Question;
        switch (lv)
        {
            case 1: icon = MessageBoxIcon.Information; break;
            case 2: icon = MessageBoxIcon.Warning; break;
            case 3: icon = MessageBoxIcon.Error; break;
            default: icon = MessageBoxIcon.None; break;
        }

        DialogResult res = MessageBox.Show(
            msg, AppName(), MessageBoxButtons.OK, icon);
    }

    /// <summary>
    /// verbose print line
    /// </summary>
    /// <param name="s">text</param>
    private static void verbose_println(string s, long lv = 1)
    {
        if (_verbose) system_println(s, lv);
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// get arguments array
    /// </summary>
    /// <returns></returns>
    private static string[] GetArgs(string line)
    {
        var arr = (new List<string>(TokenArgs(line))).Skip(1).ToArray();
        return arr;
    }

    private static IEnumerable<string> TokenArgs(string s, string prev = "")
    {
        var sb = new StringBuilder(prev);
        var seq = 0;
        var esc = '^';
        var esc2 = '^';
        var cqt = '\0';
        foreach (var c in (s + ' '))
        {
            switch (seq)
            {
                case 0:         //idle
                    {
                        if (Char.IsWhiteSpace(c)) continue;
                        else if (c == esc) seq = 2;
                        else if (c == '"') { cqt = c; seq = 3; }
                        else { sb.Append(c); seq = 1; }
                    }
                    break;
                case 1:         //normal
                    {
                        if (Char.IsWhiteSpace(c))
                        {
                            yield return sb.ToString();
                            sb.Clear();
                            seq = 0;
                        }
                        else if (c == esc) seq = 2;
                        else if (c == '"') { cqt = c; seq = 3; }
                        else sb.Append(c);
                    }
                    break;
                case 2:         //escape
                    {
                        var c1 = c;
                        sb.Append(c1);
                        seq = 1;
                    }
                    break;
                case 3:         //string
                    {
                        if (c == cqt) { seq = 1; continue; }
                        else if (c == esc2) seq = 4;
                        else sb.Append(c);
                    }
                    break;
                case 4:         //string-escape
                    {
                        var c1 = c;
                        sb.Append(c1);
                        seq = 3;
                    }
                    break;
                default:
                    break;
            }
        }
        if (sb.Length > 0) yield return sb.ToString();
    }

    /// <summary>
    /// Remove quate character from string
    /// </summary>
    /// <param name="s">target string</param>
    /// <returns></returns>
    private static string RemoveQuotation(string s)
    {
        var b_quate = false;
        var sb = new StringBuilder();
        foreach (var c in s)
        {
            if (c == '"' && !b_quate)
            {
                b_quate = true;
                continue;
            }
            b_quate = false;
            sb.Append(c);
        }
        return sb.ToString();
    }

    private static string RemoveEscape(string s)
    {
        return String.Format(s);
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// run script
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    private bool RunScript(string path)
    {
        if (!File.Exists(path)) return false;
        var enc = Encoding.GetEncoding(932);

        // filter comment(rem) line
        var lines = new List<string>();
        foreach (var line in File.ReadAllLines(path, enc))
        {
            var s = line.Trim();
            if (s.Length < 4) continue;
            if (s.Substring(0, 4).ToLower() != "rem ") continue;
            lines.Add(s.Substring(4).Trim());
        }

        foreach (var line in lines)
        {
            var s = line.Trim();
            if (s.Length < 1) continue;

            //output comment line
            if (s[0] == '#')
            {
                s = s.Substring(1).Trim();
                OutLines.Add("rem " + s);
                continue;
            }

            //output set command
            if (s[0] == '*')
            {
                s = "* " + s.Substring(1).Trim();
                Reset();
                if (ParseCommandLine(s) != 0)
                {
                    var s1 = "cmd line:" + s;
                    system_println("引数が正しくありません。\n" + s1, 3);
                    continue;
                }
                if (Run() < 0) return false;
            }
        }
        return true;
    }

    /// <summary>
    /// run command
    /// </summary>
    /// <returns>error code, 0 is ok</returns>
    private int Run()
    {
        if (Items.Count > 0)
        {
            var s = Items[0];
            Items.RemoveAt(0);
            _name = s;

            var i = s.IndexOf('=');
            if (i > 0)
            {
                _value = s.Substring(i + 1).Trim();
                _name = s.Substring(0, i).Trim();
            }
        }

        var res = 0;
        switch (Command)
        {
            case "": Cmd_Set(); break;                //set constant
            case "d": res = Cmd_Set(); break;         //set delete
            case "p": res = Cmd_Prompt(); break;      //set input string
            case "y": res = Cmd_YesNo(); break;       //set message
            case "f": res = Cmd_File(); break;        //set file
            case "g": res = Cmd_Folder(); break;      //set folder
            case "l": res = Cmd_List(); break;        //set list
            case "x": res = Cmd_Test(); break;
            default: break;
        }
        return res;
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// print version
    /// </summary>
    /// <returns></returns>
    int Cmd_Version()
    {
        Console.WriteLine(_version);
        return 0;
    }

    /// <summary>
    /// print help
    /// </summary>
    /// <returns></returns>
    private int Cmd_Help()
    {
        var msg = AppName() + @".exe {option...} [script...]

option:
  mode flag:
  --version     print version
  -h,--help     print help
  -v,--verbose  verbose mode
  -C,--console  gui mode
  -u            update mode
  -a            append mode
  -r            relative mode
  command:
  -d            delete        (set XXX=)
  -p            prompt        (set XXX={input text})
  -y            yesno         (set XXX={1:yes,0:no})
  -f            choose file   (set XXX={file_path})
  -g            choose folder (set XXX={folder_path})
  -l            list select   (set XXX={selected text})
  parameter:
  -i <path>     set input-file path
  -o <path>     set output-file path(default:.tmp/<app>.cmd)
  -c <command>  command script
  -s <path>     set script path
  -t <title>    set title
  -m <message>  set message
";
        system_println(msg, 0, true);
        return 0;
    }

    /// <summary>
    /// test command
    /// </summary>
    /// <returns></returns>
    int Cmd_Test()
    {

        Console.WriteLine("version      : " + _version);
        Console.WriteLine("verbose      : " + _verbose);
        Console.WriteLine("cui          : " + _cui);
        Console.WriteLine("append       : " + _append);
        Console.WriteLine("UpdateMode   : " + UpdateMode);
        Console.WriteLine("RelativeFlag : " + RelativeFlag);
        Console.WriteLine("CurrentPath  : " + CurrentPath);
        Console.WriteLine("InputPath    : " + InputPath);
        Console.WriteLine("OutputPath   : " + OutputPath);
        Console.WriteLine("ScriptPath   : " + ScriptPath);
        Console.WriteLine("Title        : " + Title);
        Console.WriteLine("Message      : " + Message);
        Console.WriteLine("Command      : " + Command);
        Console.WriteLine("Message      : " + Message);
        Console.WriteLine("g_cpath      : " + _g_cpath);
        Console.WriteLine("g_title      : " + _g_title);

        Console.WriteLine("name  =" + _name);
        Console.WriteLine("value =" + _value);

        Console.WriteLine("cmds.count=" + Items.Count);
        for (var i = 0; i < Items.Count; i++)
        {
            Console.WriteLine(i + ": " + Items[i]);
        }

        Console.WriteLine("outs.count=" + OutLines.Count);
        for (var i = 0; i < OutLines.Count; i++)
        {
            Console.WriteLine(i + ": " + OutLines[i]);
        }

        Console.WriteLine("base.count=" + BaseValue.Count);
        for (var i = 0; i < BaseValue.Count; i++)
        {
            var ks = BaseValue.Keys.ToArray();
            Console.WriteLine(i + ": " + ks[i] + " = " + BaseValue[ks[i]]);
        }
        return 0;
    }

    /// <summary>
    /// set string
    /// </summary>
    /// <returns></returns>
    int Cmd_Set()
    {
        // set name
        string k = _name;
        if (k.Length < 1) return 0;

        // set default value
        var v = _value;

        var s = v;
        for (var i = 0; i < Items.Count; i++)
        {
            s += " " + Items[i];
        }

        // output
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    /// <summary>
    /// set input string
    /// </summary>
    /// <returns></returns>
    int Cmd_Prompt()
    {
        // set name
        string k = _name;
        if (k.Length < 1) return 0;

        // set default value
        var v = _value;
        if (v == "")
            if (BaseValue.ContainsKey(k))
                v = BaseValue[k];

        var s = RemoveQuotation(v);
        if (s == "" && BaseValue.ContainsKey(k))
            s = BaseValue[k];

        if (_cui)
        {
            var res = Prompt(ref s);
            if (res != 0) return res;
        }
        else
        {
            Console.WriteLine("msg:" + Message);
            var res = ShowInputBox(ref s, Message, Title);
            if (res != DialogResult.OK) return -1;
        }

        // output
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    /// <summary>
    /// message
    /// </summary>
    /// <returns></returns>
    int Cmd_YesNo()
    {
        string k = _name;
        if (k.Length < 1) return 0;

        var v = (_value == "") ? "1" : _value;
        if (BaseValue.ContainsKey(k))
        {
            if (Items.Count > 0)
            {
                if (BaseValue[k] == Items[0]) v = "1";
            }
            if (Items.Count > 1)
            {
                if (BaseValue[k] == Items[1]) v = "0";
            }
        }

        var s = RemoveQuotation(Message);
        if (_cui)
        {
            if (s == "") s = _name + " ?";
            s += " [Yes/No/Cancel]";
            while (true)
            {
                Console.WriteLine(s);
                Console.Write(Title + "> ");
                var s2 = Console.ReadLine();
                if (s2 == null) return -1;
                s2 = s2.Trim().ToLower();
                if (s2.Length < 1) continue;
                switch (s2[0])
                {
                    case 'y':
                        if (Items.Count > 0) s = Items[0];
                        else s = "1";
                        break;
                    case 'n':
                        s = "";
                        if (Items.Count > 1) s = Items[1];
                        else if (Items.Count < 1) s = "0";
                        break;
                    case 'c':
                        s = "";
                        return -2;
                }
            }
        }
        else
        {
            if (s == "") s = k + " ?";
            var btn = MessageBoxDefaultButton.Button1;
            if (v == "0") btn = MessageBoxDefaultButton.Button2;

            // show dialog
            DialogResult res = MessageBox.Show(
                s, Title, MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question, btn);
            s = "";
            if (res == DialogResult.Yes)
            {
                if (Items.Count > 0) s = Items[0];
                else s = "1";
            }
            else if (res == DialogResult.No)
            {
                if (Items.Count > 1) s = Items[1];
                else s = "0";
            }
            else return -1;
        }

        // output
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    /// <summary>
    /// set path
    /// </summary>
    /// <returns></returns>
    int Cmd_File()
    {
        // set name
        string k = _name;
        if (k.Length < 1) return 0;

        // set default value
        var v = _value;
        if (v == "")
            if (BaseValue.ContainsKey(k))
                v = BaseValue[k];
        if (v == "") v = CurrentPath;

        // choose
        var res = 0;
        if (_cui)
            res = Prompt(ref v);
        else
            res = GetFile(ref v);
        if (res != 0) return res;
        var s = v;

        // output
        if (RelativeFlag) s = GetRelativePath(s, CurrentPath);
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Title = "";
        Message = "";
        return 0;
    }

    private int GetFile(ref string p)
    {
        // create dialog
        OpenFileDialog dlg = new OpenFileDialog();
        if (Message != "") dlg.Title = Message;

        // set directory and file namw
        var s = Path.GetFullPath(p);
        if (Directory.Exists(s))
        {
            dlg.InitialDirectory = s;
            s = "";
        }
        else
        {
            var d = Path.GetDirectoryName(s);
            if (d == null) d = CurrentPath;
            if (Directory.Exists(d))
                dlg.InitialDirectory = d;
            s = Path.GetFileName(s);
            dlg.FileName = s;
        }

        // set filter
        var flt = GetFileNameFilter();
        dlg.Filter = flt;
        if (s != "")
        {
            s = Path.GetExtension(s);
            if (s != "")
            {
                var i = flt.IndexOf("*" + s);
                if (i >= 0)
                {
                    s = flt.Substring(0, i);
                    i = s.Split('|').Count() / 2;
                    dlg.FilterIndex = i;
                }
            }
            s = "";
        }

        // show dialog
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            s = dlg.FileName;
        }
        dlg.Dispose();
        if (s == "") return -1;
        p = s;
        return 0;
    }

    private string GetFileNameFilter()
    {
        var k1 = "";
        var flt = "";
        foreach (var v1 in Items)
        {
            if (k1 == "")
            {
                k1 = v1;
                continue;
            }
            flt += k1 + "|" + v1 + "|";
            k1 = "";
        }
        if (flt != "")
        {
            flt += "すべてのファイル|*.*";
        }
        return flt;
    }

    /// <summary>
    /// set folder
    /// </summary>
    /// <returns></returns>
    int Cmd_Folder()
    {
        // set name
        string k = _name;
        if (k.Length < 1) return 0;

        // set default value
        var v = _value;
        if (v == "")
            if (BaseValue.ContainsKey(k))
                v = BaseValue[k];
        if (v == "") v = CurrentPath;

        // choose
        var res = 0;
        if (_cui)
            res = Prompt(ref v);
        else
            res = GetFolder(ref v);
        if (res != 0) return res;
        var s = v;

        // output
        if (RelativeFlag) s = GetRelativePath(s, CurrentPath);
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    private int GetFolder(ref string p)
    {
        FolderBrowserDialog dlg = new FolderBrowserDialog();
        dlg.ShowNewFolderButton = true;
        if (Message != "") dlg.Description = Message;
        dlg.SelectedPath = p;

        // show dialog
        var s = "";
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            s = dlg.SelectedPath;
        }
        dlg.Dispose();
        if (s == "") return -1;
        p = s;
        return 0;
    }

    /// <summary>
    /// set datalist
    /// </summary>
    /// <returns></returns>
    int Cmd_List()
    {
        // set name
        string k = _name;
        if (k.Length < 1) return 0;

        // set default value
        var v = _value;
        if (v == "")
            if (BaseValue.ContainsKey(k))
                v = BaseValue[k];

        // choose
        var res = 0;
        if (_cui)
            res = Prompt(ref v);
        else
            res = GetDataList(ref v);
        if (res != 0) return res;
        var s = v;

        // output
        if (QuotationFlag) s = "\"" + s + "\"";
        s = "set " + k + "=" + s;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    private int GetDataList(ref string v)
    {
        var dlg = new Tmm.UI.InputDialog(Message, Title, true);

        if (!Items.Contains(v)) dlg.AddListItem(v);
        foreach (var item in Items)
            dlg.AddListItem(item);
        dlg.Value = v;
        dlg.Text = Message;
        dlg.FocusList(v);

        // show dialog
        var s = "";
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            s = dlg.Value;
        }
        dlg.Dispose();
        if (s == "") return -1;
        v = s;
        return 0;
    }

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// input box
    /// </summary>
    /// <param name="s"></param>
    /// <param name="msg"></param>
    /// <param name="title"></param>
    /// <param name="w"></param>
    /// <param name="h"></param>
    /// <param name="ptn"></param>
    /// <returns></returns>
    private static DialogResult ShowInputBox(
        ref string s, string msg, string title = "Title",
        int w = 300, int h = 145, string ptn = "")
    {
        var sz = new Size(w, h);
        var m = 10;

        Form dlg = new Form();
        dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
        //dlg.AutoScaleDimensions = new SizeF(6F, 13F);
        //dlg.AutoScaleMode = AutoScaleMode.Font;
        dlg.ClientSize = sz;
        dlg.Text = title;
        dlg.MinimumSize = new Size(w, h);
        dlg.MinimizeBox = false;
        dlg.MaximizeBox = false;
        dlg.ShowIcon = false;

        Label lbl = new Label();
        lbl.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        lbl.Text = msg;
        lbl.Location = new Point(m, m);
        lbl.Width = sz.Width - 2 * m;
        lbl.Height = 3 * 20;
        dlg.Controls.Add(lbl);

        TextBox txt = new TextBox();
        txt.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
        txt.BorderStyle = BorderStyle.FixedSingle;
        txt.Size = new Size(sz.Width - 2 * m, 23);
        txt.Location = new Point(m, lbl.Location.Y + lbl.Height + m);
        txt.Text = s;
        dlg.Controls.Add(txt);

        Button ok = new Button();
        ok.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
        ok.DialogResult = DialogResult.OK;
        ok.Name = "ok";
        ok.Size = new Size(75, 23);
        ok.Text = "&OK";
        ok.Location = new Point(sz.Width - 80 - 75 - m, sz.Height - 4 * m);
        dlg.Controls.Add(ok);

        Button cancel = new Button();
        cancel.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right);
        cancel.DialogResult = DialogResult.Cancel;
        cancel.Name = "cancel";
        cancel.Size = new Size(75, 23);
        cancel.Text = "&Cancel";
        cancel.Location = new Point(sz.Width - 75 - m, sz.Height - 4 * m);
        dlg.Controls.Add(cancel);

        dlg.AcceptButton = ok;
        dlg.CancelButton = cancel;

        DialogResult res = dlg.ShowDialog();
        while (true)
        {
            if (res != DialogResult.OK) return res;
            if (ptn == "") break;
            if (Regex.Match(txt.Text, ptn).Success) break;
            res = dlg.ShowDialog();
        }
        s = txt.Text;
        return res;
    }

    /// <summary>
    /// console prompt
    /// </summary>
    /// <param name="s"></param>
    /// <returns>error code</returns>
    private int Prompt(ref string s)
    {
        var s2 = (s.Length > 0) ? " [" + s + "]" : "";
        s2 = (Message + s2).Trim();
        if (s2.Length > 0)
        {
            Console.WriteLine(s2);
        }
        Console.Write(Title + "> ");
        s2 = Console.ReadLine();
        if (s2 == null) return -2;
        s2 = s2.Trim();
        if (s2.Length > 0)
        {
            s = s2;
        }
        return 0;
    }

    /// <summary>
    /// get relative path
    /// </summary>
    /// <param name="path"></param>
    /// <param name="root"></param>
    /// <returns>path</returns>
    private static string GetRelativePath(string path, string root)
    {
        var ps = Path.GetFullPath(path.Replace("/", "\\")).Split('\\');
        var rs = Path.GetFullPath(root.Replace("/", "\\")).Split('\\');
        var n = ps.Count();
        if (n > rs.Count()) n = rs.Count();
        if (n == 0) return path;
        if (ps[0] != rs[0]) return path;
        var i = 0;
        for (; i < n; i++)
        {
            if (ps[i] != rs[i]) break;
        }
        var res = "";
        for (var j = i; j < ps.Count(); j++)
        {
            res = res + "\\" + ps[j];
        }
        if (i < rs.Count())
        {
            for (; i < rs.Count() - 1; i++)
            {
                res = "\\.." + res;
            }
            res = "." + res;
        }
        res = "." + res;
        return res;
    }

    void Load(string path)
    {
        if (path == "") return;
        if (!File.Exists(path)) return;

        var enc = Encoding.GetEncoding(932);
        foreach (var line in File.ReadLines(path))
        {
            var ss = line.Split(new char[] { ' ' }, 2);
            if (ss.Length != 2) continue;
            ss = ss[1].Split(new char[] { '=' }, 2);
            if (ss.Length != 2) continue;
            var k = ss[0];
            var v = ss[1];
            if (BaseValue.ContainsKey(k)) BaseValue.Remove(k);
            BaseValue.Add(k, v);
        }
    }

    /// <summary>
    /// save to output path
    /// </summary>
    /// <param name="path"></param>
    void Save(string path)
    {
        if (OutLines.Count == 0) return;

        if (path == "")
        {
            foreach (var line in OutLines)
            {
                Console.WriteLine(line);
            }
            return;
        }

        if (!_append)
        {
            if (File.Exists(path)) File.Delete(path);
        }

        var enc = Encoding.GetEncoding(932);
        foreach (var line in OutLines)
        {
            var s = line + Environment.NewLine;
            File.AppendAllText(path, s, enc);
        }
    }
}

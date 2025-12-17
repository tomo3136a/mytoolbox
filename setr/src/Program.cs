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
            var s1 = "command line:" + line;
            system_println("引数が正しくありません。\n" + s1, 3);
            Environment.ExitCode = -1;
            return;
        }
        app.SetGlobal();

        // output file
        var dst = app.Load();

        var src = app.InputPath;
        if (src != "")
        {
            // test source file
            if (!File.Exists(src))
            {
                var s1 = src;
                system_println("入力ファイルがありません。\n" + s1, 3);
                Environment.ExitCode = -1;
                return;
            }

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
        else
        {
            if (app.Run() < 0)
            {
                Environment.ExitCode = -1;
                return;
            }
        }

        // save output line data
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

    //////////////////////////////////////////////////////////////////////////
    /// 

    /// <summary>
    /// load initialize file
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    string Load(string path = "")
    {
        var f = path;
        if (f == "") f = OutputPath;
        if (!File.Exists(f)) return "";

        var enc = Encoding.GetEncoding(932);
        foreach (var line in File.ReadAllLines(f, enc))
        {
            var ss = line.Split(new char[] { ' ' }, 2);
            if (ss.Length != 2) continue;
            ss = ss[1].Split(new char[] { '=' }, 2);
            if (ss.Length != 2) continue;
            var k = ss[0];
            var v = ss[1];
            v = ApplyEnvironment(v, BaseValue);
            if (BaseValue.ContainsKey(k)) BaseValue.Remove(k);
            BaseValue.Add(k, v);
        }
        return f;
    }

    /// <summary>
    /// save to output path
    /// </summary>
    /// <param name="path"></param>
    void Save(string path = "")
    {
        if (OutLines.Count == 0) return;

        var f = path;
        if (f == "") f = OutputPath;

        if (f == "")
        {
            foreach (var line in OutLines)
            {
                Console.WriteLine(line);
            }
            return;
        }

        if (!_append)
        {
            if (File.Exists(f)) File.Delete(f);
        }

        var enc = Encoding.GetEncoding(932);
        foreach (var line in OutLines)
        {
            var s = line + Environment.NewLine;
            File.AppendAllText(f, s, enc);
        }
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
        var ptn = "^[\\s@]*rem\\s+([#*])(.*)$";
        var re1 = new Regex(ptn, RegexOptions.IgnoreCase);
        var re2 = new Regex("[^\\^](\\^\\^)*\\^$");

        var buf = "";
        foreach (var line in File.ReadAllLines(path, enc))
        {
            var m = re1.Match(line);
            if (!m.Success)
            {
                buf = "";
                continue;
            }
            var t = m.Groups[1].Value;
            var s = m.Groups[2].Value;

            //comment line
            if (t == "#")
            {
                buf = "";
                OutLines.Add("rem " + s);
                continue;
            }

            //continus line
            m = re2.Match(s);
            if (m.Success)
            {
                buf += s.Substring(0, s.Length - 1);
                continue;
            }
            s = ApplyEnvironment(buf + s);
            buf = "";

            //set command
            Reset();
            if (ParseCommandLine(s, 0) != 0)
            {
                var s1 = "cmd line:" + s;
                system_println("引数が正しくありません。\n" + s1, 3);
                continue;
            }
            if (Run() < 0) return false;
        }
        return true;
    }

    /// <summary>
    /// Commandline parser
    /// </summary>
    /// <param name="line">commandline</param>
    /// <returns>error code</returns>
    private int ParseCommandLine(string line, int skip = 1)
    {
        var atk = new Tmm.ArgsTokener();

        var re = new Regex(@"^(--?|/)(\w+)(?:(-|\+)?|=(.*))?$");
        var opt_flg = false;
        var opt = "";
        var res = 0;
        foreach (var arg in atk.Tokens(line).Skip(skip))
        {
            var s = arg.Trim();
            if (s.Length == 0) continue;
            if (!opt_flg)
            {
                //analyze option
                var m = re.Match(s);
                if (!m.Success)
                {
                    s = RemoveQuotation(s);
                    Items.Add(s);
                    continue;
                }
                opt = m.Groups[2].Value;

                //long option name
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

                //short option name
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
                    case "b": Command = opt; continue; //message
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

                s = m.Groups[4].Value;
                opt_flg = true;
            }
            if (!opt_flg || s.Length == 0) continue;
            s = RemoveQuotation(s);
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

    /// <summary>
    /// run command
    /// </summary>
    /// <returns>error code, 0 is ok</returns>
    private int Run()
    {
        var res = 0;

        foreach (var s in Items)
        {
            _name = s;

            var i = s.IndexOf('=');
            if (i > 0)
            {
                _value = s.Substring(i + 1).Trim();
                _name = s.Substring(0, i).Trim();
            }

            switch (Command)
            {
                case "": Cmd_Set(); break;                //set constant
                case "d": res = Cmd_Set(); break;         //set delete
                case "p": res = Cmd_Prompt(); break;      //set input string
                case "b": res = Cmd_MsgBox(); break;      //set message
                case "y": res = Cmd_YesNo(); break;       //set yesyno
                case "f": res = Cmd_File(); break;        //set file
                case "g": res = Cmd_Folder(); break;      //set folder
                case "l": res = Cmd_List(); break;        //set list
                case "x": res = Cmd_Test(); break;
                default: break;
            }

            if (res != 0) break;
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
  -q            quotation mode
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
  -c <path>     set current-directory path
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
        var k = _name;

        // set value
        var v = _value;
        v = v.Trim();

        // output
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
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
        var k = _name;

        // set value
        var v = _value;
        v = RemoveQuotation(v);
        if (v == "" && BaseValue.ContainsKey(k))
            v = BaseValue[k];

        var msg = Message;
        if (msg == "") msg = k + "?";

        if (_cui)
        {
            var res = Prompt(ref v);
            if (res != 0) return res;
        }
        else
        {
            var res = ShowInputBox(ref v, msg, Title);
            if (res != DialogResult.OK) return -1;
        }

        // output
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
        OutLines.Add(s);
        Message = "";
        return 0;
    }

    /// <summary>
    /// message dialog
    /// </summary>
    /// <returns></returns>
    int Cmd_MsgBox()
    {
        var msg = "";
        foreach (var s in Items)
            msg += " " + s;
        msg = msg.Trim();

        if (_cui)
            Console.WriteLine(msg);
        else
            MessageBox.Show(
                msg, Title, MessageBoxButtons.OK,
                MessageBoxIcon.Question);

        return 1;
    }

    /// <summary>
    /// message
    /// </summary>
    /// <returns></returns>
    int Cmd_YesNo()
    {
        // set name
        var k = _name;

        var msg = RemoveQuotation(Message);
        if (msg == "") msg = k + "?";
        if (_cui) msg += " [Yes/No/Cancel]";

        // set value
        var v0 = "0";
        var v1 = "1";
        if (Items.Count > 1) { v0 = ""; v1 = Items[1]; }
        if (Items.Count > 2) { v0 = Items[2]; }
        var v = _value;
        v = RemoveQuotation(v);
        if (v == "" && BaseValue.ContainsKey(k))
            v = BaseValue[k];
        if (v == "") v = v1;

        if (_cui)
        {
            while (true)
            {
                Console.WriteLine(msg);
                Console.Write(Title + "> ");
                var s2 = Console.ReadLine();
                if (s2 == null) return -1;
                s2 = s2.Trim().ToLower();
                if (s2.Length < 1) continue;
                switch (v[0])
                {
                    case 'y': v = v1; break;
                    case 'n': v = v0; break;
                    case 'c': return -2;
                }
                break;
            }
        }
        else
        {
            var btn = MessageBoxDefaultButton.Button1;
            if (v != v1) btn = MessageBoxDefaultButton.Button2;

            // show dialog
            switch (MessageBox.Show(
                msg, Title, MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question, btn))
            {
                case DialogResult.Yes: v = v1; break;
                case DialogResult.No: v = v0; break;
                default: return -1;
            }
        }

        // output
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
        OutLines.Add(s);
        Message = "";
        return 1;
    }

    /// <summary>
    /// set path
    /// </summary>
    /// <returns></returns>
    int Cmd_File()
    {
        // set name
        string k = _name;

        // set value
        var v = _value;
        v = v.Trim();

        if (v == "")
            if (BaseValue.ContainsKey(k))
                v = BaseValue[k];
        if (v == "") v = CurrentPath;

        // choose
        var res = 0;
        if (_cui)
            res = Prompt(ref v);
        else
            res = GetFile(ref v, Items.Skip(1).ToList());
        if (res != 0) return res;

        // output
        if (RelativeFlag) v = GetRelativePath(v, CurrentPath);
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
        OutLines.Add(s);
        Title = "";
        Message = "";
        return 1;
    }

    private int GetFile(ref string p, List<string> flts)
    {
        // create dialog
        OpenFileDialog dlg = new OpenFileDialog();
        if (Message != "") dlg.Title = Message;

        // set directory and file namw
        var s = Path.GetFullPath(p);
        if (Directory.Exists(s))
        {
            Console.WriteLine("InitialDirectory:" + s + ":");
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
            dlg.FileNames.Append(s);
            dlg.FileName = s;
        }

        // set filter
        var flt = GetFileNameFilter(flts);
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

    private static string GetFileNameFilter(List<string> items)
    {
        var k1 = "";
        var flt = "";
        foreach (var v1 in items)
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

        // set value
        var v = _value;
        v = v.Trim();

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

        // output
        if (RelativeFlag) v = GetRelativePath(v, CurrentPath);
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
        OutLines.Add(s);
        Message = "";
        return 1;
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

        // set value
        var v = _value;
        v = v.Trim();

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

        // output
        if (QuotationFlag) v = "\"" + v + "\"";
        var s = "set " + k + "=" + v;
        OutLines.Add(s);
        Message = "";
        return 1;
    }

    private int GetDataList(ref string v)
    {
        var dlg = new Tmm.UI.InputDialog(Message, Title, true);

        if (v != "")
            if (!Items.Contains(v)) dlg.AddListItem(v);
        foreach (var item in Items.Skip(1))
            dlg.AddListItem(item);
        dlg.Value = v;
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

    private static string ApplyEnvironment(string s, Dictionary<string, string> dic = null)
    {
        var sb = new StringBuilder();
        var kw = "";
        var seq = 0;
        var ret = 0;
        foreach (var c in s)
        {
            switch (seq)
            {
                case 0:
                    if (c == '^') seq = 2;
                    else if (c == '%') { ret = 0; seq = 3; }
                    else { sb.Append(c); }
                    if (c == '"') seq = 1;
                    break;
                case 1:
                    if (c == '%') { ret = 1; seq = 3; break; }
                    sb.Append(c);
                    if (c == '"') seq = 0;
                    break;
                case 2:
                    sb.Append(c);
                    seq = 0;
                    break;
                case 3:
                    if (Char.IsLetterOrDigit(c) || c == '_')
                    {
                        kw += c;
                        break;
                    }
                    if (c == '%')
                        kw = GetEnvironmentVariable(kw, dic);
                    else
                        kw = "%" + kw + c;
                    sb.Append(kw);
                    kw = "";
                    seq = ret;
                    break;
            }
        }
        return sb.ToString();
    }

    private static string GetEnvironmentVariable(string kw, Dictionary<string, string> dic)
    {
        var v = kw;
        if (dic != null)
        {
            if (dic.ContainsKey(kw))
            {
                v = dic[kw];
            }
            else
                v = Environment.GetEnvironmentVariable(kw);
        }
        else
            v = Environment.GetEnvironmentVariable(kw);
        return v;
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
}

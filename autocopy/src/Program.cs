using System;
using System.Windows.Forms;
using System.Threading;

namespace AutoCopy
{
    static class Program
    {
        private static Service _serv = Service.GetCurrentService();

        private static string JobName = "job" + Config.GetName();
        private static Mutex mutexObject = new Mutex();

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            OperatingSystem os = Environment.OSVersion;
            if ((os.Platform == PlatformID.Win32NT) && (os.Version.Major >= 5))
            {
                JobName = @"Global\" + JobName;
            }

            using (mutexObject = new Mutex(false, JobName))
            {
                if (!mutexObject.WaitOne(100, true))
                {
                    if (!mutexObject.WaitOne(100000, true))
                    {
                        MessageBox.Show("タイムアウトしました。\n" + JobName,
                            Config.GetName(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        mutexObject.Close();
                        return;
                    }
                }
                AppMain(args);
                mutexObject.ReleaseMutex();
            }
            mutexObject.Close();
        }

        static void AppMain(string[] args)
        {
            if (!CommandLineParse(args)) return;
            //_serv.SetServiceMode();  //test code
            if (_serv.IsService())
            {
                _serv.Initialize();
                _serv.Action();
                return;
            }
#if true
            //Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
#else
            ApplicationConfiguration.Initialize();
#endif
            Application.Run(new CopyForm());
        }

        private static bool WaitProcessTimeout()
        {
            OperatingSystem os = Environment.OSVersion;
            //if ((os.Platform == PlatformID.Win32NT) && (os.Version.Major >= 5))
            //{
            JobName = @"Global\" + JobName;
            //}

            using (mutexObject = new Mutex(false, JobName))
            {
                if (!mutexObject.WaitOne(100, true))
                {
                    if (!mutexObject.WaitOne(100000, true))
                    {
                        MessageBox.Show("タイムアウトしました。\n" + JobName,
                            Config.GetName(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        mutexObject.Close();
                        return true;
                    }
                }
                mutexObject.ReleaseMutex();
            }
            mutexObject.Close();
            return false;
        }

        /// <summary>
        /// コマンドライン解析
        /// -b  バックグラウンド動作
        /// -c  コンフィグパス指定
        /// *** アクションパス
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static bool CommandLineParse(string[] args)
        {
            var seq = 0;
            foreach (var arg in args)
            {
                switch (seq)
                {
                    case 0:
                        if (arg.StartsWith("-") || arg.StartsWith("/"))
                        {
                            switch (arg.Substring(1).ToLower())
                            {
                                case "s": _serv.SetServiceMode(); continue;
                                case "h": Help(); return false;
                                case "g": _serv.SetGuiMode(); continue;
                                case "c": seq = 1; continue;    //config file
                                case "f": seq = 2; continue;    //load entry file
                            }
                            UI.ShowMessage("オプション " + arg + " は未サポートです。");
                            return false;
                        }
                        _serv.AddEntry(arg);
                        break;
                    case 1:
                        _serv.SetConfigPath(arg);
                        seq = 0;
                        break;
                    case 2:
                        _serv.LoadEntriesFile(arg);
                        seq = 0;
                        break;
                    default:
                        seq = -1;
                        break;
                }
            }
            return (seq == 0);
        }
        static void Help()
        {
            UI.AddMessage("Usage:");
            UI.AddMessage("> " + Config.GetName() + " <options>... <entry>...");
            UI.AddMessage("");
            UI.AddMessage("-s        : set service mode");
            UI.AddMessage("-g        : set gui mode");
            UI.AddMessage("-h        : help");
            UI.AddMessage("-c <path> : configure file path");
            UI.AddMessage("-f <path> : load entry fille list");
            UI.ShowMessage();
        }
    }
}

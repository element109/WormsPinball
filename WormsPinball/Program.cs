namespace WormsPinball
{
    using Microsoft.Win32;
    using System.Diagnostics;

    internal static class Program
    {
        static string tmpDDraw;
        static string tmpConfig;

        [STAThread]
        static void Main(string[] args)
        {
#if DOTNET48
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
#else
            ApplicationConfiguration.Initialize();
#endif
            Application.ThreadException += Application_ThreadException;

            bool isFirstInstance;
            using (var mtx = new Mutex(true, "19F5A7D5-D672-4279-8AF4-3416D8576445", out isFirstInstance))
            {
                if (isFirstInstance)
                {
                    string exe = @"C:\Program Files (x86)\Steam\steamapps\common\Worms Pinball\WormsPinball.exe";
                    string InstallDirectory = @"C:\Program Files (x86)\Steam\steamapps\common\Worms Pinball";
                    var key = Registry.CurrentUser.OpenSubKey(@"Software\Valve\Steam", false);
                    if (key != null)
                    {
                        var steamDir = key.GetValue("SteamPath", @"C:\Program Files (x86)\Steam").ToString();
                        InstallDirectory = Path.Combine(steamDir, @"steamapps\common\Worms Pinball");
                        exe = Path.Combine(InstallDirectory, "WormsPinball.exe");
                        key.Close();
                    }

                    if (!Directory.Exists(InstallDirectory))
                    {
                        InstallDirectory = Environment.CurrentDirectory;
                        exe = Path.Combine(InstallDirectory, "WormsPinball.exe");
                    }

                    var wormsPinball = new FileInfo(exe);
                    if (wormsPinball.Exists)
                    {
                        try
                        {
                            var dDraw = Path.Combine(wormsPinball.DirectoryName, "DDraw.dll");
                            tmpDDraw = Path.Combine(Path.GetTempPath(), "DDraw.dll");
                            if (File.Exists(dDraw))
                            {
                                File.Copy(dDraw, tmpDDraw, true);

                                var dgVoodoo = Path.Combine(wormsPinball.DirectoryName, "dgVoodoo.conf");
                                tmpConfig = Path.Combine(Path.GetTempPath(), "dgVoodoo.conf");
                                if (File.Exists(dgVoodoo))
                                    File.Copy(dgVoodoo, tmpConfig, true);

                                var info = new ProcessStartInfo();
                                info.WorkingDirectory = wormsPinball.DirectoryName;
                                info.FileName = wormsPinball.FullName;
                                if (args.Length > 0)
                                    info.Arguments = String.Join(' ', args);

                                using (var wp = Process.Start(info))
                                {
                                    wp.EnableRaisingEvents = true;
                                    wp.Exited += Process_Exited;

                                    Application.Run();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Message(ex.ToString(), ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        static void Process_Exited(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(tmpDDraw))
                    File.Delete(tmpDDraw);

                if (File.Exists(tmpConfig))
                    File.Delete(tmpConfig);
            }
            catch (Exception ex)
            {
                Message(ex.ToString(), ex.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Exit();
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            Message(e.Exception.ToString(), e.Exception.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Exit();
        }

        static void Exit()
        {
            Application.Exit();
        }

        static DialogResult Message(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(message, title, buttons, icon, MessageBoxDefaultButton.Button1);
        }
    }
}
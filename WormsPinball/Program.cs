namespace WormsPinball
{
    using Microsoft.Win32;
    using System.Diagnostics;
    using System.Management;

    internal static class Program
    {
        static bool foundSoundConfig;
        static bool soundMute;
        static bool musicEnabled = true;
        static float soundVolume = 100;
        static int soundDelay = 5000;

        [STAThread]
        static void Main(string[] args)
        {
#if DOTNET48
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
#else
            ApplicationConfiguration.Initialize();
#endif
            Application.ThreadException += (sender, e) =>
            {
                Message(e.Exception.ToString(), e.Exception.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            };
            
            using (var mtx = new Mutex(true, "19F5A7D5-D672-4279-8AF4-3416D8576445", out bool isFirstInstance))
            {
                if (isFirstInstance)
                {
                    string InstallDirectory = @"C:\Program Files (x86)\Steam\steamapps\common\Worms Pinball";
                    var key = Registry.CurrentUser.OpenSubKey(@"Software\Valve\Steam", false);
                    if (key != null)
                    {
                        var steamDir = key.GetValue("SteamPath", @"C:\Program Files (x86)\Steam").ToString();
                        InstallDirectory = Path.Combine(steamDir, @"steamapps\common\Worms Pinball");
                        key.Close();
                    }

                    if (!Directory.Exists(InstallDirectory))
                        InstallDirectory = Environment.CurrentDirectory;

                    var sndConfig = Path.Combine(InstallDirectory, "sound.cfg");
                    if (File.Exists(sndConfig))
                    {
                        var lines = File.ReadAllLines(sndConfig);
                        Parse(lines);
                        foundSoundConfig = true;
                    }

                    string exe;
                    if (musicEnabled)
                        exe = Path.Combine(InstallDirectory, "WormsPinball.exe");
                    else
                        exe = Path.Combine(InstallDirectory, "WormsPinballNoMusic.exe");

                    var wormsPinball = new FileInfo(exe);
                    if (wormsPinball.Exists)
                    {
                        string tmpDDraw = null, tmpConfig = null;
                        if (musicEnabled)
                        {
                            var tmpPath = Path.GetTempPath();
                            var dDraw = Path.Combine(wormsPinball.DirectoryName, "DDraw.dll");
                            tmpDDraw = Path.Combine(tmpPath, "DDraw.dll");
                            if (File.Exists(dDraw))
                                File.Copy(dDraw, tmpDDraw, true);

                            var dgVoodoo = Path.Combine(wormsPinball.DirectoryName, "dgVoodoo.conf");
                            tmpConfig = Path.Combine(tmpPath, "dgVoodoo.conf");
                            if (File.Exists(dgVoodoo))
                                File.Copy(dgVoodoo, tmpConfig, true);
                        }

                        var info = new ProcessStartInfo();
                        info.WorkingDirectory = wormsPinball.DirectoryName;
                        info.FileName = wormsPinball.FullName;
                        if (args.Length > 0)
                            info.Arguments = string.Join(' ', args);

                        using (var wp = Process.Start(info))
                        {
                            wp.EnableRaisingEvents = true;
                            wp.Exited += (sender, e) =>
                            {
                                if (musicEnabled)
                                {
                                    try
                                    {
                                        if (File.Exists(tmpDDraw))
                                            File.Delete(tmpDDraw);

                                        if (File.Exists(tmpConfig))
                                            File.Delete(tmpConfig);
                                    }
                                    catch { }
                                }

                                Application.Exit();
                            };

                            if (foundSoundConfig)
                            {
                                double delay = soundDelay;
                                var delayTimer = new System.Timers.Timer(delay)
                                {
                                    AutoReset = false
                                };
                                delayTimer.Elapsed += (sender, e) =>
                                {
                                    int id = musicEnabled ? GetChildProcessID(wp.Id) : wp.Id;
                                    if (id > -1)
                                    {
                                        var app = Windows.AudioManager.FindApplication(id);
                                        if (app != null)
                                        {
                                            app.Volume = soundVolume;
                                            app.Mute = soundMute;
                                            app.Dispose();
                                        }
                                    }

                                    delayTimer.Dispose();
                                };

                                delayTimer.Start();
                            }

                            Application.Run();
                        }
                    }
                }
            }
        }

        static void Parse(string[] lines)
        {
            foreach (var line in lines)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    if (!line.StartsWith('#'))
                    {
                        if (line.StartsWith("volume", StringComparison.OrdinalIgnoreCase))
                        {
                            var volume = line.Split('=', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            if (volume.Length > 1)
                            {
                                if (float.TryParse(volume[1], out float result))
                                    soundVolume = result;
                            }
                        }
                        else if (line.StartsWith("mute", StringComparison.OrdinalIgnoreCase))
                        {
                            var mute = line.Split('=', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            if (mute.Length > 1)
                            {
                                if (bool.TryParse(mute[1], out bool result))
                                    soundMute = result;
                            }
                        }
                        else if (line.StartsWith("music", StringComparison.OrdinalIgnoreCase))
                        {
                            var music = line.Split('=', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            if (music.Length > 1)
                            {
                                if (bool.TryParse(music[1], out bool result))
                                    musicEnabled = result;
                            }
                        }
                        else if (line.StartsWith("delay", StringComparison.OrdinalIgnoreCase))
                        {
                            var delay = line.Split('=', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            if (delay.Length > 1)
                            {
                                if (int.TryParse(delay[1], out int result))
                                    if (result > 0)
                                        soundDelay = result;
                            }
                        }
                    }
                }
            }
        }

        static int GetChildProcessID(int parentId)
        {
            using var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Process WHERE ParentProcessId={parentId}");
            using ManagementObjectCollection collection = searcher.Get();
            if (collection.Count > 0)
            {
                foreach (var item in collection)
                {
                    if (item["ProcessId"] is uint result)
                        return Convert.ToInt32(result);
                }
            }
            return -1;
        }

        static DialogResult Message(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(message, title, buttons, icon, MessageBoxDefaultButton.Button1);
        }
    }
}
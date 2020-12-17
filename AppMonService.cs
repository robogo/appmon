using System;
using System.ComponentModel;
using System.Configuration;
using System.Configuration.Install;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Globalization;
using System.Text;

namespace appmon
{
    public partial class AppMonService : ServiceBase
    {
        const int QueryCommand = 200;
        const int AddTimeCommand = 201;
        const int RemoveTimeCommand = 202;
        const int EmailCommand = 203;
        public const string Name = "appmon";
        string appName;
        string appKeyword;
        string message;
        string email;
        Timer timer;
        int weekday;
        int weekend;
        int bonus;
        TimeSpan interval;
        bool running;
        List<TimeRange> activeTimes;
        DateTime lastTime;
        int usedSeconds;
        Notification notification;
        bool debug;

        public AppMonService()
        {
            InitializeComponent();
            this.ServiceName = Name;
        }

        protected override void OnStart(string[] args)
        {
            this.Init();
            this.timer = new Timer(Monitor, this, TimeSpan.FromSeconds(5), Timeout.InfiniteTimeSpan);
        }

        protected override void OnStop()
        {
            lock (this.activeTimes)
            {
                this.Write(this.activeTimes);
            }

            this.timer?.Dispose();
        }

        protected override void OnCustomCommand(int command)
        {
            switch (command)
            {
                case QueryCommand:
                    int quota = IsWeekend(DateTime.Now) ? this.weekend : this.weekday;
                    int used = this.usedSeconds / 60;
                    int remaining = Math.Max(0, quota + this.bonus - used);
                    var sb = new StringBuilder(1024);
                    sb.AppendLine($"{DateTime.Now:yyyy-MM-dd}: Allowed:{quota} Bonus:{this.bonus} Used:{used} Remaining:{remaining}");
                    lock (this.activeTimes)
                    {
                        foreach (var tr in this.activeTimes)
                        {
                            DateTime start = new DateTime(tr.StartTicks, DateTimeKind.Local);
                            DateTime end = new DateTime(tr.EndTicks, DateTimeKind.Local);
                            sb.AppendFormat("{0:HH:mm:ss}-{1:HH:mm:ss}  {2}\r\n", start, end, (int)(end - start).TotalMinutes);
                        }
                    }
                    this.notification.Show("Info", sb.ToString(), 0);
                    break;
                case AddTimeCommand:
                    this.bonus += 10;
                    break;
                case RemoveTimeCommand:
                    this.bonus -= 10;
                    break;
                case EmailCommand:
                    this.Email();
                    break;
                default:
                    break;
            }

            base.OnCustomCommand(command);
        }

        public void Init()
        {
            this.appName = ConfigurationManager.AppSettings.Get("Name", s => s);
            this.appKeyword = ConfigurationManager.AppSettings.Get("Keyword", s => s, null);
            this.message = ConfigurationManager.AppSettings.Get("Message", s => s, "Stop");
            this.email = ConfigurationManager.AppSettings.Get("Email", s => s, null);
            this.weekday = ConfigurationManager.AppSettings.Get("Weekday", s => int.Parse(s), 120);
            this.weekend = ConfigurationManager.AppSettings.Get("Weekend", s => int.Parse(s), 180);
            this.interval = TimeSpan.FromSeconds(ConfigurationManager.AppSettings.Get("Interval", s => int.Parse(s), 60));
            int sesionId = ConfigurationManager.AppSettings.Get("SessionId", s => int.Parse(s), -1);
            this.notification = new Notification(sesionId);
            this.debug = ConfigurationManager.AppSettings.Get("Debug", s => s == "true", false);
            this.activeTimes = this.Read();
            this.usedSeconds = (int)(this.activeTimes.Select(tr => tr.EndTicks - tr.StartTicks).Sum() / 10000000);
            this.lastTime = DateTime.Now;
            this.running = false;
            if (this.debug)
            {
                EventLog.WriteEntry(Name, "Service initialized", EventLogEntryType.Information);
            }
        }

        public static void Monitor(object state)
        {
            var thisPtr = (AppMonService)state;

            try
            {
                var now = DateTime.Now;
                var last = thisPtr.lastTime;
                if (last.Year < now.Year || last.Month < now.Month || last.Day < now.Day)
                {
                    int quota = IsWeekend(now) ? thisPtr.weekend : thisPtr.weekday;
                    int used = thisPtr.usedSeconds / 60;
                    int remaining = Math.Max(0, quota + thisPtr.bonus - used);
                    thisPtr.bonus = Math.Min(remaining, 60);

                    lock (thisPtr.activeTimes)
                    {
                        thisPtr.activeTimes.Clear();
                    }

                    thisPtr.lastTime = now;
                    thisPtr.running = false;
                    thisPtr.usedSeconds = 0;
                    if (thisPtr.debug)
                    {
                        EventLog.WriteEntry(Name, $"Data reset on a new day. Roll over: {thisPtr.bonus}", EventLogEntryType.Information);
                    }
                }
                else if (now.Subtract(last + thisPtr.interval) > TimeSpan.FromSeconds(30))
                {
                    thisPtr.lastTime = now;
                    thisPtr.running = false;
                    if (thisPtr.debug)
                    {
                        EventLog.WriteEntry(Name, "Reset due to missing updates", EventLogEntryType.Information);
                    }
                }

                string query = $"SELECT Name, CommandLine, ProcessId, Caption, ExecutablePath FROM Win32_Process WHERE Name = '{thisPtr.appName}'";
                string wmiScope = @"\\.\root\cimv2";
                var searcher = new ManagementObjectSearcher(wmiScope, query);
                bool match = false;
                uint procId = 0;
                foreach (ManagementObject mo in searcher.Get())
                {
                    match = true;
                    procId = (uint)mo["ProcessId"];
                    if (thisPtr.debug)
                    {
                        EventLog.WriteEntry(Name, "Found process " + string.Join(",", mo.Properties.Cast<PropertyData>().Select(o => $"{o.Name}:{o.Value}")), EventLogEntryType.Information);
                    }

                    if (!string.IsNullOrEmpty(thisPtr.appKeyword))
                    {
                        match = false;
                        var cl = (string)mo["CommandLine"];
                        if (!string.IsNullOrEmpty(cl) && cl.Contains(thisPtr.appKeyword))
                        {
                            match = true;
                            break;
                        }

                        var path = (string)mo["ExecutablePath"];
                        if (!string.IsNullOrEmpty(path) && path.Contains(thisPtr.appKeyword))
                        {
                            match = true;
                            break;
                        }

                        if (thisPtr.debug)
                        {
                            EventLog.WriteEntry(Name, $"Keyword not matched cl={cl} path={path}", EventLogEntryType.Information);
                        }
                    }
                }

                if (match)
                {
                    if (thisPtr.running)
                    {
                        lock (thisPtr.activeTimes)
                        {
                            int lastIndex = thisPtr.activeTimes.Count - 1;
                            TimeRange tr = thisPtr.activeTimes[lastIndex];
                            thisPtr.activeTimes[lastIndex] = new TimeRange() { StartTicks = tr.StartTicks, EndTicks = now.Ticks };
                        }

                        thisPtr.usedSeconds += (int)now.Subtract(thisPtr.lastTime).TotalSeconds;
                        int quota = IsWeekend(now) ? thisPtr.weekend : thisPtr.weekday;
                        int minutes = thisPtr.usedSeconds / 60;
                        if (thisPtr.debug)
                        {
                            EventLog.WriteEntry(Name, $"Used={minutes} Quota={quota} Bonus={thisPtr.bonus}", EventLogEntryType.Information);
                        }

                        if (minutes > quota + thisPtr.bonus)
                        {
                            if (thisPtr.debug)
                            {
                                EventLog.WriteEntry(Name, $"Display warning to session {thisPtr.notification.SessionId}", EventLogEntryType.Information);
                            }

                            thisPtr.notification.Show(
                                "Warning",
                                $"{thisPtr.message}.\r\n{quota + thisPtr.bonus} minutes allowed. {minutes} minutes used.",
                                10);
                        }
                    }
                    else
                    {
                        lock (thisPtr.activeTimes)
                        {
                            thisPtr.activeTimes.Add(new TimeRange() { StartTicks = now.Ticks, EndTicks = now.Ticks });
                        }

                        thisPtr.running = true;
                        if (thisPtr.debug)
                        {
                            EventLog.WriteEntry(Name, "Found process. Start counting", EventLogEntryType.Information);
                        }
                    }

                    thisPtr.lastTime = now;
                }
                else
                {
                    if (thisPtr.running)
                    {
                        thisPtr.running = false;
                        if (thisPtr.debug)
                        {
                            EventLog.WriteEntry(Name, "Process stopped. Stop counting", EventLogEntryType.Information);
                        }
                    }
                    else
                    {
                        if (thisPtr.debug)
                        {
                            EventLog.WriteEntry(Name, "No process found", EventLogEntryType.Information);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                EventLog.WriteEntry(Name, exception.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                thisPtr.timer?.Change(thisPtr.interval, Timeout.InfiniteTimeSpan);
            }
        }

        static bool IsWeekend(DateTime d)
        {
            return d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday;
        }

        static string GetDataDir()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Name);
        }

        static string GetDataPath()
        {
            return Path.Combine(GetDataDir(), "usage.data");
        }

        List<TimeRange> Read()
        {
            var list = new List<TimeRange>();
            string path = GetDataPath();
            if (File.Exists(path))
            {
                string[] lines = File.ReadAllLines(path);
                DateTime dt = DateTime.ParseExact(lines[0], "yyyy-MM-dd", CultureInfo.InvariantCulture);
                DateTime today = DateTime.Now;
                if (dt.Year == today.Year && dt.Month == today.Month && dt.Day == today.Day && !string.IsNullOrEmpty(lines[1]))
                {
                    foreach (string item in lines[1].Split(','))
                    {
                        string[] range = item.Split('-');
                        long start = long.Parse(range[0]);
                        long end = long.Parse(range[1]);
                        list.Add(new TimeRange() { StartTicks = start, EndTicks = end });
                    }
                }

                if (lines.Length > 2)
                {
                    this.bonus = int.Parse(lines[2]);
                }
            }

            return list;
        }

        void Write(List<TimeRange> ranges)
        {
            var now = DateTime.Now;
            var last = this.lastTime;
            if (last.Year == now.Year && last.Month == now.Month && last.Day == now.Day)
            {
                string line1 = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string line2 = string.Join(",", ranges.Select(tr => $"{tr.StartTicks}-{tr.EndTicks}").ToArray());
                string line3 = this.bonus.ToString();
                string dir = GetDataDir();
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                string path = GetDataPath();
                File.WriteAllLines(path, new[] { line1, line2, line3 });
            }
        }

        void Email()
        {
            if (string.IsNullOrEmpty(this.email))
            {
                return;
            }
        }
    }

    struct TimeRange
    {
        public long StartTicks;
        public long EndTicks;
    }

    [RunInstaller(true)]
    public class MyWindowsServiceInstaller : Installer
    {
        public MyWindowsServiceInstaller()
        {
            var processInstaller = new ServiceProcessInstaller();
            var serviceInstaller = new ServiceInstaller();

            //set the privileges
            processInstaller.Account = ServiceAccount.LocalSystem;

            serviceInstaller.DisplayName = "app mon";
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            serviceInstaller.ServiceName = AppMonService.Name;
            this.Installers.Add(processInstaller);
            this.Installers.Add(serviceInstaller);
        }
    }

    sealed class Notification
    {
        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern bool WTSSendMessage(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.I4)] int SessionId,
            String pTitle,
            [MarshalAs(UnmanagedType.U4)] int TitleLength,
            String pMessage,
            [MarshalAs(UnmanagedType.U4)] int MessageLength,
            [MarshalAs(UnmanagedType.U4)] int Style,
            [MarshalAs(UnmanagedType.U4)] int Timeout,
            [MarshalAs(UnmanagedType.U4)] out int pResponse,
            bool bWait);

        bool sending;

        public Notification(int sessionId)
        {
            this.SessionId = sessionId;
        }

        public int SessionId { get; private set; }

        public async void Show(string title, string message, int timeout)
        {
            if (this.sending)
            {
                return;
            }

            this.sending = true;
            try
            {
                await Task.Run(() =>
                {
                    int sessionId = this.SessionId > 0 ? this.SessionId : 1;
                    while (sessionId < 8)
                    {
                        bool ret = WTSSendMessage(IntPtr.Zero, sessionId, title, title.Length,
                            message, message.Length, 0, timeout, out int resp, true);
                        if (ret)
                        {
                            if (resp == 32000 && this.SessionId < 0)
                            {
                                // IDTIMEOUT
                                sessionId++;
                                continue;
                            }

                            if (this.SessionId < 0)
                            {
                                this.SessionId = sessionId;
                            }

                            break;
                        }

                        int err = Marshal.GetLastWin32Error();
                        EventLog.WriteEntry(AppMonService.Name, $"Send message to session {sessionId} failed with error {err}", EventLogEntryType.Error);

                        if (this.SessionId > 0 || err != 2)
                        {
                            break;
                        }

                        sessionId++;
                    }
                });
            }
            finally
            {
                this.sending = false;
            }
        }
    }

    public static class Win32Extension
    {
        [Flags]
        public enum ThreadAccess : int
        {
            TERMINATE = (0x0001),
            SUSPEND_RESUME = (0x0002),
            GET_CONTEXT = (0x0008),
            SET_CONTEXT = (0x0010),
            SET_INFORMATION = (0x0020),
            QUERY_INFORMATION = (0x0040),
            SET_THREAD_TOKEN = (0x0080),
            IMPERSONATE = (0x0100),
            DIRECT_IMPERSONATION = (0x0200)
        }

        [DllImport("kernel32.dll")]
        static extern IntPtr OpenThread(ThreadAccess dwDesiredAccess, bool bInheritHandle, uint dwThreadId);
        [DllImport("kernel32.dll")]
        static extern uint SuspendThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        static extern int ResumeThread(IntPtr hThread);

        public static T Get<T>(this NameValueCollection collection, string key, Func<string, T> func)
        {
            string value = collection[key];
            if (string.IsNullOrEmpty(value))
            {
                throw new KeyNotFoundException(key);
            }

            return func(value);
        }

        public static T Get<T>(this NameValueCollection collection, string key, Func<string, T> func, T def)
        {
            string value = collection[key];
            if (string.IsNullOrEmpty(value))
            {
                return def;
            }

            return func(value);
        }

        public static void Suspend(this Process process)
        {
            foreach (ProcessThread thread in process.Threads)
            {
                var pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)thread.Id);
                if (pOpenThread == IntPtr.Zero)
                {
                    break;
                }

                SuspendThread(pOpenThread);
            }
        }

        public static void Resume(this Process process)
        {
            foreach (ProcessThread thread in process.Threads)
            {
                var pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)thread.Id);
                if (pOpenThread == IntPtr.Zero)
                {
                    break;
                }

                ResumeThread(pOpenThread);
            }
        }
    }
}

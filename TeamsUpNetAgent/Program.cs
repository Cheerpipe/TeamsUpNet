using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace TeamsUpNetAgent
{

    internal class Util
    {

        public static void minimizeMemory()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle,
                (UIntPtr)0xFFFFFFFF, (UIntPtr)0xFFFFFFFF);
        }

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process,
            UIntPtr minimumWorkingSetSize, UIntPtr maximumWorkingSetSize);
    }

    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            int frequency;
            Util.minimizeMemory();
            while (true)
            {
                frequency = 2000;
                if (IsTeamsVisible())
                {
                    if (!IsTeamsUpUp())
                        Process.Start("teamsupnet.exe");
                }
                else
                {
                    var hora = DateTime.Now.Hour;
                    var day = DateTime.Now.DayOfWeek;

                    if (hora >= 19 || day == DayOfWeek.Saturday || day == DayOfWeek.Sunday)
                    {
                        KillProcessByName("dropbox.exe");
                        KillProcessByName("dropboxupdate.exe");
                        KillProcessByName("teamsupnet.exe");
                        return;
                    }
                    if (IsTeamsUpUp())
                        KillProcessByName("teamsupnet.exe");
                    if (!IsTeamsUp)
                        frequency = 60000;

                }
                Thread.Sleep(frequency);
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);
        static bool IsTeamsUp = false;
        public static bool IsTeamsVisible()
        {
            bool _IsIconic = false;
            try
            {
                Process p = Process.GetProcessesByName("teams").Where(pp => pp.MainWindowHandle != IntPtr.Zero).First();
                _IsIconic = IsIconic(p.MainWindowHandle);
                p.Dispose();
                IsTeamsUp = true;
                return !_IsIconic;
            }
            catch
            {
                IsTeamsUp = false;
                return false;
            }
        }

        public static bool IsTeamsUpUp()
        {
            return Process.GetProcessesByName("teamsupnet").Length > 0;
        }

        public static void KillProcessByName(string processName)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = @"C:\Windows\System32\taskkill.exe";
            cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            cmd.StartInfo.Arguments = string.Format(@"/f /im {0}", processName);
            cmd.Start();
            cmd.Dispose();
        }
    }
}

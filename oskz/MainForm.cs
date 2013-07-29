using System;
using System.Drawing;
using System.Windows.Forms;
using MouseKeyboardActivityMonitor.WinApi;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MouseKeyboardActivityMonitor;

namespace MouseKeyboardActivityMonitor.OSKZ
{
    public partial class MainForm : Form
    {
        private readonly KeyboardHookListener m_KeyboardHookManager;

        public MainForm()
        {
            GimmeTray();

            // init
            InitializeComponent();

            // Block key
            m_KeyboardHookManager = new KeyboardHookListener(new GlobalHooker());
            m_KeyboardHookManager.Enabled = true;

            // Show OSK if need
            //GimmeOSK();
        }

        // Tray ===============================================================================================================

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;

        private void GimmeTray()
        {
            // Create a simple tray menu with only one item.
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);

            // Create a tray icon.
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Windows System Key Blocker";
            trayIcon.Icon = new Icon(SystemIcons.Shield, 40, 40);

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
        }

        private void OnExit(object sender, EventArgs e)
        {
            Console.WriteLine(" ! OSK exited.");
            m_KeyboardHookManager.Dispose();
            Application.Exit();
        }

        // OSK ===============================================================================================================

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd,
            UInt32 Msg,
            IntPtr wParam,
            IntPtr lParam);
        private const UInt32 WM_SYSCOMMAND = 0x112;
        private const UInt32 SC_RESTORE = 0xf120;

        private const string OnScreenKeyboardExe = "osk.exe";

        private void GimmeOSK()
        {
            Process[] ps = Process.GetProcessesByName(
            Path.GetFileNameWithoutExtension(OnScreenKeyboardExe));

            if (ps.Length == 0)
            {
                // we must start osk from an MTA thread
                if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                {
                    ThreadStart start = new ThreadStart(StartOSK);
                    Thread thread = new Thread(start);
                    thread.SetApartmentState(ApartmentState.MTA);
                    thread.Start();
                    thread.Join();
                }
                else
                {
                    StartOSK();
                }

                ps = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(OnScreenKeyboardExe));
                Process p = ps[0];
                p.EnableRaisingEvents = true;
                p.Exited += new EventHandler(myProcess_HasExited);
            }
            else
            {
                // there might be a race condition if the process terminated 
                // meanwhile -> proper exception handling should be added
                //
                SendMessage(ps[0].MainWindowHandle,
                    WM_SYSCOMMAND, new IntPtr(SC_RESTORE), new IntPtr(0));
            }
        }

        private void myProcess_HasExited(object sender, System.EventArgs e)
        {
            Console.WriteLine(" ! OSK exited.");
            m_KeyboardHookManager.Dispose();
            Application.Exit();
        }

        private void StartOSK()
        {
            Console.WriteLine(" * StartOSK");

            IntPtr ptr = new IntPtr(); ;
            bool sucessfullyDisabledWow64Redirect = false;

            // Disable x64 directory virtualization if we're on x64,
            // otherwise keyboard launch will fail.
            if (System.Environment.Is64BitOperatingSystem)
            {
                sucessfullyDisabledWow64Redirect =
                    Wow64DisableWow64FsRedirection(ref ptr);
            }

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = OnScreenKeyboardExe;
            // We must use ShellExecute to start osk from the current thread
            // with psi.UseShellExecute = false the CreateProcessWithLogon API 
            // would be used which handles process creation on a separate thread 
            // where the above call to Wow64DisableWow64FsRedirection would not 
            // have any effect.
            //
            psi.UseShellExecute = false;

            // Start the process.
            Process p = Process.Start(psi);

            // Re-enable directory virtualisation if it was disabled.
            if (System.Environment.Is64BitOperatingSystem)
                if (sucessfullyDisabledWow64Redirect)
                    Wow64RevertWow64FsRedirection(ptr);

            // Wait for the window to finish loading.
            p.WaitForInputIdle();

            p.EnableRaisingEvents = true;
            p.Exited += new EventHandler(myProcess_HasExited);

            Console.WriteLine(" ! OSK Ready");
        }

        // init ===============================================================================================================

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // MainForm
            // 
            this.CausesValidation = false;
            this.ClientSize = new System.Drawing.Size(116, 22);
            this.ControlBox = false;
            this.Name = "MainForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.ResumeLayout(false);
            this.Visible = false;
        }
    }
}
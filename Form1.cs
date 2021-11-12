using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ListApp.Properties;
using CefSharp.WinForms;
using CefSharp;
using System.Threading.Tasks;

namespace ListApp
{
    public partial class Form1 : Form
    {
        private const int gripOffset = 16;
        private const int menuBarOffset = 32;

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private ChromiumWebBrowser browser;

        public Form1()
        {
            InitializeComponent();

            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Info", OnInfo);
            trayMenu.MenuItems.Add("Exit", OnExit);
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Notepad";
            trayIcon.Icon = new Icon("notepad.ico");
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            trayIcon.MouseClick += new MouseEventHandler(trayIcon_Click);

            this.Location = Settings.Default.WinLoc;
            this.Size = Settings.Default.WinSize;
            this.Opacity = Settings.Default.Opacity;

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            browser = new ChromiumWebBrowser("https://www.icloud.com/notes");
            browser.FrameLoadEnd += new EventHandler<CefSharp.FrameLoadEndEventArgs>(FrameLoadEnd);
            browser.KeyboardHandler = new KeyboardHandler(this);
            browser.Dock = DockStyle.Fill;
            this.Controls.Add(browser);
        }

        private void FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            Console.WriteLine(Settings.Default.Zoom);
            browser.SetZoomLevel(Settings.Default.Zoom);
        }

        private void trayIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                return;
            }
            if (this.WindowState == FormWindowState.Normal)
            {
                WindowState = FormWindowState.Minimized;
                Visible = false;
                ShowInTaskbar = false;
            }
            else
            {
                Visible = true;
                ShowInTaskbar = false;
                WindowState = FormWindowState.Normal;
                Activate();
            }
        }

        private void OnInfo(object sender, EventArgs e)
        {
            MessageBox.Show("Ctrl + OemMinus: Opacity down\nCtrl + OemPlus: Opacity up");
        }

        private void OnExit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
   
            Application.Exit();
            System.Environment.Exit(1);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (systemShutdown)
            {
                CleanUp();
                Application.Exit();
                System.Environment.Exit(1);
            }

            CleanUp();

            browser.Dispose();
            Cef.Shutdown();
        }
        private void CleanUp()
        {
            if (this.WindowState != FormWindowState.Minimized)
            {
                Settings.Default.WinLoc = this.Location;
                Settings.Default.WinSize = this.Size;
            }

            Settings.Default.Opacity = this.Opacity;
            double temp = Settings.Default.Zoom;
            if (!browser.IsDisposed)
            {
                Task<double> task = browser.GetZoomLevelAsync();
                task.Wait();
                temp = task.Result;
            }
            Console.WriteLine(temp);
            Settings.Default.Zoom = temp;
            Settings.Default.Save();
        }

        private static int WM_QUERYENDSESSION = 0x11;
        private static bool systemShutdown = false;
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x20)
            {  // Trap WM_SETCUROR
                if ((m.LParam.ToInt32() & 0xffff) == 2)
                { // Trap HTCAPTION
                    Cursor.Current = Cursors.Hand;
                    m.Result = (IntPtr)1;  // Processed
                    return;
                }
            }
            if (m.Msg == 0x84)
            {  // Trap WM_NCHITTEST
                Point pos = new Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);

                if (pos.Y < menuBarOffset)
                {
                    m.Result = (IntPtr)2;  // HTCAPTION
                    return;
                }

                if (pos.X >= this.ClientSize.Width - gripOffset && pos.Y >= this.ClientSize.Height - gripOffset)
                {
                    m.Result = (IntPtr)17; // HTBOTTOMRIGHT
                    return;
                }

                if (pos.X <= gripOffset && pos.Y >= this.ClientSize.Height - gripOffset)
                {
                    m.Result = (IntPtr)16; // HTBOTTOMLEFT
                    return;
                }
                if (pos.X <= gripOffset)
                {
                    m.Result = (IntPtr)10; // HTLEFT
                    return;
                }
                if (pos.X >= this.ClientSize.Width - gripOffset)
                {
                    m.Result = (IntPtr)11; // HTRIGHT
                    return;
                }
                if (pos.Y >= this.ClientSize.Height - gripOffset)
                {
                    m.Result = (IntPtr)15; //HTBOTTOM
                    return;
                }

            }

            if (m.Msg == WM_QUERYENDSESSION)
            {
                systemShutdown = true;
            }

            base.WndProc(ref m);
        }
    }

    public class KeyboardHandler : IKeyboardHandler
    {
        private Form1 form1;

        public KeyboardHandler(Form1 form1)
        {
            this.form1 = form1;
        }

        public bool OnKeyEvent(IWebBrowser browserControl, IBrowser browser, KeyType type, int windowsKeyCode, int nativeKeyCode, CefEventFlags modifiers, bool isSystemKey)
        {
            return false;
        }
        public bool OnPreKeyEvent(IWebBrowser browserControl, IBrowser browser, KeyType type, int windowsKeyCode, int nativeKeyCode, CefEventFlags modifiers, bool isSystemKey, ref bool isKeyboardShortcut)
        {
            if (windowsKeyCode == 48 && modifiers == CefEventFlags.ControlDown)
            {
                Task<double> task = browser.GetZoomLevelAsync();
                task.ContinueWith(previous =>
                {
                    if (previous.IsCompleted)
                    {
                        double currentLevel = previous.Result;
                        browser.SetZoomLevel(currentLevel + 0.05);
                    }
                    else
                    {
                        throw new InvalidOperationException("Unexpected failure of calling CEF->GetZoomLevelAsync", previous.Exception);
                    }
                }, TaskContinuationOptions.ExecuteSynchronously);
                return true;
            }
            if (windowsKeyCode == 57 && modifiers == CefEventFlags.ControlDown)
            {
                Task<double> task = browser.GetZoomLevelAsync();
                task.ContinueWith(previous =>
                {
                    if (previous.IsCompleted)
                    {
                        double currentLevel = previous.Result;
                        browser.SetZoomLevel(currentLevel - 0.05);
                    }
                    else
                    {
                        throw new InvalidOperationException("Unexpected failure of calling CEF->GetZoomLevelAsync", previous.Exception);
                    }
                }, TaskContinuationOptions.ExecuteSynchronously);
                return true;
            }
            if (windowsKeyCode == 187 && modifiers == CefEventFlags.ControlDown)
            {
                form1.BeginInvoke((Action)(() => { form1.Opacity += 0.05; }));
                return true;
            }
            if (windowsKeyCode == 189 && modifiers == CefEventFlags.ControlDown)
            {
                form1.BeginInvoke((Action)(() => { form1.Opacity -= 0.05; }));
                return true;
            }
            return false;
        }
    }
}

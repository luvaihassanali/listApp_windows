using System;
using System.Drawing;
using System.Windows.Forms;
using ListApp.Properties;
using CefSharp.WinForms;
using CefSharp;
using System.Timers;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ListApp
{
    public partial class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        //Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        private const int MOUSEEVENTF_WHEEL = 0x0800;

        private const int gripOffset = 16;
        private const int menuBarOffset = 32;

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        ContextMenuStrip contextMenu;
        private ChromiumWebBrowser browser;
        private bool firstCall = true;
        private System.Timers.Timer t;
        private bool onNotesPage = false;
        private bool closing = false;
        string pageSource;

        public Form1()
        {
            InitializeComponent();
            //InitializeTextbox();
            InitializeBrowser();

            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Info", OnInfo);
            trayMenu.MenuItems.Add("Exit", OnExit);
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Notepad";
            trayIcon.Icon = new Icon("notepad.ico");
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            trayIcon.MouseClick += new MouseEventHandler(trayIcon_Click);

            /*if(!File.Exists("notes.rtf"))
            {
                this.richTextBox1.SaveFile("notes.rtf", RichTextBoxStreamType.RichText);
            } 
            else
            {
                this.richTextBox1.LoadFile("notes.rtf", RichTextBoxStreamType.RichText);
            }*/

            this.Location = Settings.Default.WinLoc;
            this.Size = Settings.Default.WinSize;
            this.Opacity = Settings.Default.Opacity;

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyItem = new ToolStripMenuItem("Copy");
            copyItem.Image = Properties.Resources.copy;
            copyItem.Click += DoCopy;
            contextMenu.Items.Add(copyItem);

            ToolStripMenuItem pasteItem = new ToolStripMenuItem("Paste");
            pasteItem.Image = Properties.Resources.paste;
            pasteItem.Click += DoPaste;
            contextMenu.Items.Add(pasteItem);

            ToolStripMenuItem cutItem = new ToolStripMenuItem("Cut");
            cutItem.Image = Properties.Resources.cut;
            cutItem.Click += DoCut;
            contextMenu.Items.Add(cutItem);
        }

        private void InitializeBrowser()
        {
            if (!Cef.IsInitialized) // Check before init
            {
                CefSettings settings = new CefSettings();
                //settings.CefCommandLineArgs.Add("disable-web-security");
                Cef.Initialize(settings, performDependencyCheck: true, browserProcessHandler: null);
            }

            browser = new ChromiumWebBrowser("https://www.icloud.com/notes");
            browser.FrameLoadEnd += new EventHandler<CefSharp.FrameLoadEndEventArgs>(FrameLoadEnd);
            //browser.KeyboardHandler = new KeyboardHandler(this);
            browser.Dock = DockStyle.Fill;
            this.Controls.Add(browser);
        }

        private void FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            if (e.Url.Contains("https://www.icloud.com/applications/notes3/current/en-us/index.html?rootDomain=www"))
            {
                onNotesPage = true;
                t.Interval = 3000;
                t.Start();
            }
            if (e.Frame.IsMain)
            {
                //browser.SetZoomLevel(Settings.Default.Zoom);
                if (e.Url.Contains("https://www.icloud.com/notes"))
                {
                    t = new System.Timers.Timer();
                    t.Interval = 2500; // In milliseconds
                    t.AutoReset = true;
                    t.Elapsed += new ElapsedEventHandler(TimerElapsed);
                    t.Start();

                }
            }
        }
        public void DoMouseClick()
        {
            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            //browser.ShowDevTools();
            IBrowser iBrowser = browser.GetBrowser();
            //List<string> frameNames = iBrowser.GetFrameNames();
            IFrame iFrame = iBrowser.GetFrame("Widget"); //aid-auth-widget //aid-auth-widget-iFrame

            if (closing)
            {
                closing = false;
                t.Stop();
                browser.Invoke(new MethodInvoker(delegate { 
                    browser.Visible = false;
                    this.Controls.Remove(browser);
                    browser.Dispose();
                }));
                //InitializeTextbox();
                here
                return;
            }
            if (onNotesPage)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    this.Cursor = new Cursor(Cursor.Current.Handle);
                    this.BringToFront();
                    this.Activate();
                    int left = this.DesktopLocation.X;
                    int top = this.DesktopLocation.Y;
                    Cursor.Position = new Point(left + 200, top + 200);
                    DoMouseClick();
                    SendKeys.SendWait("^a"); 

                    Task t1 = Task.Run(GetSource);
                    t1.Wait();

                    //Console.WriteLine(pageSource);
                }));

                t.Stop();
                t.Interval = 1000;
                t.Start();
                closing = true;
                return;
            }

            if (firstCall)
            {
                iFrame.ExecuteJavaScriptAsync("document.getElementById('account_name_text_field').focus();");
                iFrame.ExecuteJavaScriptAsync("document.getElementById('account_name_text_field').value=" + '\'' + "luvaihassanali@gmail" + '\'');
                SendKeys.SendWait(".com");
                SendKeys.SendWait("{ENTER}");
                t.Interval = 1000;
                firstCall = false;
            }
            else
            {
                string readText = File.ReadAllText("secret.txt");
                Console.WriteLine(readText);
                //To-do: wrap in await
                iFrame.ExecuteJavaScriptAsync("document.getElementById('password_text_field').focus();");
                iFrame.ExecuteJavaScriptAsync("document.getElementById('password_text_field').value=" + '\'' + readText + '\'');
                SendKeys.SendWait("0");
                SendKeys.SendWait("{ENTER}");
                t.Stop();
            }
        }

        private async void GetSource()
        {
            IBrowser iBrowser = browser.GetBrowser();
            List<string> frameNames = iBrowser.GetFrameNames();
            IFrame iFrame = iBrowser.GetFrame(frameNames[2]);
            pageSource = await iFrame.GetSourceAsync();
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
                richTextBox1.Focus();
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                Activate();
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //this.richTextBox1.SaveFile("notes.rtf", RichTextBoxStreamType.RichText);
            Settings.Default.WinLoc = this.Location;
            Settings.Default.WinSize = this.Size;
            Settings.Default.Opacity = this.Opacity;
            Settings.Default.Save();
        }

        private void OnInfo(object sender, EventArgs e)
        {
            MessageBox.Show("Ctrl + b: Bold/Unbold \nCtrl + s: Opacity down\nCtrl + d: Opacity up");
        }

        private void OnExit(object sender, EventArgs e)
        {
            //this.richTextBox1.SaveFile("notes.rtf", RichTextBoxStreamType.RichText);

            if (this.WindowState != FormWindowState.Minimized)
            {
                Settings.Default.WinLoc = this.Location;
                Settings.Default.WinSize = this.Size;
                Settings.Default.Opacity = this.Opacity;
            }

            trayIcon.Visible = false;
            trayIcon.Dispose();

            Settings.Default.Save();
            Application.Exit();
            System.Environment.Exit(1);
        }

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
            base.WndProc(ref m);
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.B)
            {
                if (richTextBox1.SelectionFont.Bold)
                {
                    richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Regular);
                }
                else
                {
                    richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Bold);
                }
            }

            if (e.Control && e.KeyCode == Keys.H)
            {
                if (richTextBox1.SelectionBackColor == Color.Yellow)
                {
                    richTextBox1.SelectionBackColor = Color.PaleGoldenrod;
                }
                else
                {
                    richTextBox1.SelectionBackColor = Color.Yellow;
                }
            }

            if (e.Control && (e.KeyCode == Keys.Oemplus))
            {
                this.Opacity += 0.05;
                e.SuppressKeyPress = true;
            }

            if (e.Control && e.KeyCode == Keys.OemMinus)
            {
                this.Opacity -= 0.05;
                e.SuppressKeyPress = true;
            }
        }
        private void richTextBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            contextMenu.Show(this, this.PointToClient(MousePosition));
        }

        private void DoCopy(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox1.SelectedText);
        }

        private void DoPaste(object sender, EventArgs e)
        {
            //DataFormats.Format myFormat = DataFormats.GetFormat(DataFormats.Rtf);

            // if (richTextBox1.CanPaste(myFormat))
            {
                richTextBox1.Paste();
            }
        }

        private void DoCut(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }
    }
}

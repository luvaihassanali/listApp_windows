using System;
using System.Drawing;
using System.Windows.Forms;
using ListApp.Properties;
using CefSharp.WinForms;
using CefSharp;
using System.Timers;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ListApp
{
    public partial class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        private const int MOUSEEVENTF_WHEEL = 0x0800;

        private const int gripOffset = 16;
        private const int menuBarOffset = 32;

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private ContextMenuStrip contextMenu;
        private Word.Application wordApp;
        private ChromiumWebBrowser browser;
        private System.Timers.Timer setupTimer;

        private bool onNotesPage = false;
        private bool foundFirstHeaderSymbol = true;
        private bool importing = true;
        private bool exporting = false;
        private bool exit = false;
        private bool shutdown = false;
        private bool systemShutdown = false;

        public Form1()
        {
            InitializeComponent();
            InitializeBrowser();
            InitContextMenuAndTrayIcon();
        }

        #region General form functions

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!shutdown)
            {
                e.Cancel = true;
                Save();
                Console.WriteLine("saving");
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            Settings.Default.WinLoc = this.Location;
            Settings.Default.WinSize = this.Size;
            Settings.Default.Opacity = this.Opacity;
            Settings.Default.Save();
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
                //richTextBox1.Focus();
                //richTextBox1.SelectionStart = richTextBox1.Text.Length;
                Activate();
            }
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

        public void DoMouseClick()
        {
            //Call the imported function with the cursor's current position
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }

        #endregion

        #region Cef Browser functions

        private void InitializeBrowser()
        {
            if (!Cef.IsInitialized) // Check before init
            {
                CefSettings settings = new CefSettings();
                //settings.LogSeverity = LogSeverity.Verbose;
                settings.CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\\Cache");
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
                if (importing)
                {
                    onNotesPage = true;
                    setupTimer.Interval = 2500;
                    Console.WriteLine("frameLoadEnd importing");
                    setupTimer.Start();

                }
                else
                {
                    Console.WriteLine("frameLoadEnd exporting");
                    exporting = true;
                    setupTimer.Start();
                }
            }
            if (e.Frame.IsMain)
            {
                //browser.SetZoomLevel(Settings.Default.Zoom);
                if (e.Url.Contains("https://www.icloud.com/notes"))
                {
                    setupTimer = new System.Timers.Timer();
                    setupTimer.Interval = 2500; // In milliseconds
                    setupTimer.AutoReset = true;
                    setupTimer.Elapsed += new ElapsedEventHandler(TimerElapsed);
                    Console.WriteLine("FrameLoadEnd main");
                    //setupTimer.Start();

                }
            }
        }

        #endregion

        #region Timer 

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            //browser.ShowDevTools();
            IBrowser iBrowser = browser.GetBrowser();
            //List<string> frameNames = iBrowser.GetFrameNames();

            if (exit)
            {
                #region exit app

                //wait for final save
                System.Threading.Thread.Sleep(8000);

                Console.WriteLine("Word quit export");
                wordApp.Quit(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                setupTimer.Stop();
                setupTimer.AutoReset = false;

                this.Invoke(new MethodInvoker(delegate
                {
                    browser.Dispose();
                    Console.WriteLine("CEF shutdown");
                    Cef.Shutdown();
                }));


                if (systemShutdown)
                {
                    System.Diagnostics.Process.Start("shutdown", "/s /t 0");
                }

                shutdown = true;
                System.Environment.Exit(1);

                #endregion
            }

            if (exporting)
            {
                #region export notes to cloud

                exporting = false;

                this.Invoke(new MethodInvoker(delegate
                {
                    this.Cursor = new Cursor(Cursor.Current.Handle);
                    this.BringToFront();
                    this.Activate();
                    int left = this.DesktopLocation.X;
                    int top = this.DesktopLocation.Y;

                    //To-do: find a way to not use mouse
                    Cursor.Position = new Point(left + 200, top + 200);
                    DoMouseClick();

                    SendKeys.SendWait("^a");
                    SendKeys.SendWait("{BACKSPACE}");

                    System.Threading.Thread.Sleep(1000);

                    Word.Document currDoc = wordApp.ActiveDocument;

                    //To-do: Highlight
                    //To-do: Fix bolding errors with symbols
                    for (int i = 1; i < currDoc.Characters.Count; i++)
                    {
                        Word.Range currCharRange = currDoc.Characters[i];
                        string currCharStr = currCharRange.Text;
                        char currChar = currCharStr.ToCharArray()[0];
                        int currNumRep = currChar - '0';
                        Console.WriteLine(i + " " + currCharStr + " " + currChar + " " + currNumRep);

                        if (currNumRep == -35)
                        {
                            //if (i == currDoc.Characters.Count - 1 || i == currDoc.Characters.Count - 2) continue;
                            System.Threading.Thread.Sleep(5);
                            SendKeys.Send("{ENTER}");
                            System.Threading.Thread.Sleep(5);
                            continue;
                        }

                        System.Threading.Thread.Sleep(5);
                        currCharRange.Copy();
                        System.Threading.Thread.Sleep(5);
                        browser.GetFocusedFrame().Paste();
                        System.Threading.Thread.Sleep(5);
                    }
                }));

                Console.WriteLine("end of export");
                setupTimer.Interval = 2000;
                setupTimer.Start();
                exporting = false;
                exit = true;

                return;

                #endregion
            }

            if (onNotesPage)
            {
                if (!importing)
                {
                    setupTimer.Stop();
                    return;
                }

                #region import notes to app

                this.Invoke(new MethodInvoker(delegate
                {
                    this.Cursor = new Cursor(Cursor.Current.Handle);
                    this.BringToFront();
                    this.Activate();
                    int left = this.DesktopLocation.X;
                    int top = this.DesktopLocation.Y;
                    //To-do: find a way to not use mouse
                    Cursor.Position = new Point(left + 200, top + 200);
                    DoMouseClick();
                    SendKeys.SendWait("^a");
                    browser.GetFocusedFrame().Copy();
                    Console.WriteLine("import copy");
                }));

                setupTimer.Stop();

                this.Invoke(new MethodInvoker(delegate
                {
                    InitializeTextbox();
                }));

                InitializeWord(true);

                return;

                #endregion

            }
        }

        #endregion

        #region Microsoft Word

        private void InitializeWord(bool import)
        {
            wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document currDoc = wordApp.Documents.Add();

            wordApp.Selection.Paste();
            currDoc.Range().ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            object noSpacingStyle = "No Spacing";
            currDoc.Range().set_Style(noSpacingStyle);
            currDoc.Range().ParagraphFormat.SpaceBefore = 0.0f;
            currDoc.Range().ParagraphFormat.SpaceAfter = 0.0f;
            currDoc.Range().Font.Name = "Arial";
            currDoc.Range().Font.Size = 12;

            currDoc.ActiveWindow.Selection.WholeStory();
            currDoc.ActiveWindow.Selection.Copy();
            Console.WriteLine("currdoc copy");

            if (import)
            {
                richTextBox1.Invoke(new MethodInvoker(delegate
                {
                    richTextBox1.Focus();
                    richTextBox1.SelectAll();
                    richTextBox1.Paste();

                    browser.Invoke(new MethodInvoker(delegate
                    {
                        browser.Visible = false;
                        this.Controls.Remove(browser);
                        browser.Dispose();
                    }));

                }));

                Console.WriteLine("Word quit import");
                wordApp.Quit(false);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            else
            {
                richTextBox1.Visible = false;
                this.Controls.Remove(richTextBox1);
                richTextBox1.Dispose();
                importing = false;
                InitializeBrowser();
            }

            if (import)
            {
                richTextBox1.Invoke(new MethodInvoker(delegate
                {
                    FixSpacing(richTextBox1, ">", "•");
                }));
            }
            else
            {
                importing = false;
            }
        }

        private void FixSpacing(RichTextBox rtb, String target, String subTarget)
        {
            if (target == "")
            {
                return;
            }

            //Replace > except first with \n> but change to * to avoid infinite loop. Then change * back to >
            int headerStart = rtb.SelectionStart, headerStartIndex = 0, headerIndex;
            while ((headerIndex = rtb.Text.IndexOf(target, headerStartIndex)) != -1)
            {
                if (foundFirstHeaderSymbol)
                {
                    foundFirstHeaderSymbol = false;
                    headerStartIndex = headerIndex + target.Length;
                    continue;
                }
                rtb.SelectionStart = headerIndex;
                rtb.SelectionLength = 1;
                rtb.SelectedText = "\n*";
                headerStartIndex = headerIndex + target.Length;

            }

            //change * to >
            target = "*";
            headerStart = rtb.SelectionStart;
            headerStartIndex = 0;
            headerIndex = 0;
            while ((headerIndex = rtb.Text.IndexOf(target, headerStartIndex)) != -1)
            {
                rtb.SelectionStart = headerIndex;
                rtb.SelectionLength = 1;
                rtb.SelectedText = ">";
                headerStartIndex = headerIndex + target.Length;

            }

            //Replace all • with tab + •
            int subHeaderStart = rtb.SelectionStart, subHeaderStartIndex = 0, subHeaderIndex;
            while ((subHeaderIndex = rtb.Text.IndexOf(subTarget, subHeaderStartIndex)) != -1)
            {
                rtb.SelectionStart = subHeaderIndex;
                rtb.SelectionLength = 1;
                rtb.SelectedText = " *";
                subHeaderStartIndex = subHeaderIndex + subTarget.Length;
            }

            target = "*";
            headerStart = rtb.SelectionStart;
            headerStartIndex = 0;
            headerIndex = 0;
            while ((headerIndex = rtb.Text.IndexOf(target, headerStartIndex)) != -1)
            {
                rtb.SelectionStart = headerIndex;
                rtb.SelectionLength = 1;
                rtb.SelectedText = subTarget;
                headerStartIndex = headerIndex + target.Length;
            }

            int extraSpaceStart = rtb.SelectionStart, extraSpaceStartIndex = 0, extraSpaceIndex;
            while ((extraSpaceIndex = rtb.Text.IndexOf(" \n", extraSpaceStartIndex)) != -1)
            {
                rtb.SelectionStart = extraSpaceIndex;
                rtb.SelectionLength = 3;
                if (extraSpaceIndex == rtb.Text.IndexOf(" \n\n"))
                {
                    rtb.SelectedText = "\n\n";
                }
                else
                {
                    rtb.SelectedText = "\n";
                }
                extraSpaceStartIndex = extraSpaceIndex + target.Length;
            }

            int endOfLineIndex = rtb.Text.Length - 2;
            rtb.SelectionStart = endOfLineIndex;
            rtb.SelectionLength = 2;
            rtb.SelectedText = "";
        }

        #endregion

        #region Context menu functions 

        private void OnShutdown(object sender, EventArgs e)
        {
            systemShutdown = true;
            Save();
        }

        private void OnSave(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()
        {
            richTextBox1.Focus();
            richTextBox1.SelectAll();
            richTextBox1.Copy();

            InitializeWord(false);
        }

        private void OnInfo(object sender, EventArgs e)
        {
            MessageBox.Show("Ctrl + b: Bold/Unbold \nCtrl + s: Opacity down\nCtrl + d: Opacity up");
        }

        private void OnExit(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Minimized)
            {
                Settings.Default.WinLoc = this.Location;
                Settings.Default.WinSize = this.Size;
                Settings.Default.Opacity = this.Opacity;
            }

            trayIcon.Visible = false;
            trayIcon.Dispose();

            Settings.Default.Save();
            this.Close();
        }

        private void InitContextMenuAndTrayIcon()
        {
            //To-do: icons
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Shutdown", OnShutdown);
            //trayMenu.MenuItems.Add("Save", OnSave);
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

        #endregion

        #region Rich text box functions

        //To-do: add • •
        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                richTextBox1.SelectedText = "  • ";
                e.SuppressKeyPress = true;
            }

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
            richTextBox1.Copy();
        }

        private void DoPaste(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void DoCut(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        #endregion

    }
}

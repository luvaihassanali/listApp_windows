using System;
using System.Drawing;
using System.Windows.Forms;
using ListApp.Properties;
using CefSharp.WinForms;
using CefSharp;
using System.Timers;
using System.IO;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

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
        private ContextMenuStrip contextMenu;
        private ChromiumWebBrowser browser;
        private System.Timers.Timer setupTimer;
        private bool firstTimerTickFlag = true;
        private bool onNotesPage = false;
        private bool foundFirstHeaderSymbol = true;
        private bool importing = true;
        private string pageSource;

        public Form1()
        {
            InitializeComponent();
            InitializeBrowser();

            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Save", OnSave);
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Console.WriteLine("CEF shutdown");
            Cef.Shutdown();
        }

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
            browser.KeyboardHandler = new KeyboardHandler(this);
            browser.Dock = DockStyle.Fill;
            this.Controls.Add(browser);
        }

        private void FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            if (e.Url.Contains("https://www.icloud.com/applications/notes3/current/en-us/index.html?rootDomain=www"))
            {
                if(importing)
                {
                    onNotesPage = true;
                    setupTimer.Interval = 5000;
                    Console.WriteLine("frameLoadEnd importing");
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
                    Console.WriteLine("FrameLoadend main");
                    setupTimer.Start();

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

            if (onNotesPage)
            {
                if(!importing)
                {
                    setupTimer.Stop();
                    return;
                }

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
                }));

                setupTimer.Stop();
                browser.Invoke(new MethodInvoker(delegate
                {
                    browser.Visible = false;
                    this.Controls.Remove(browser);
                    browser.Dispose();
                    //Cef.Shutdown();
                }));

                //To-do use diagnostic tool to find memory usage
                this.Invoke(new MethodInvoker(delegate
                {
                    InitializeTextbox();
                }));

                InitializeWord(true);
                return;
            }

            //To-do: test green text see how it shows up on apple
            IFrame iFrame = iBrowser.GetFrame("Widget"); //aid-auth-widget //aid-auth-widget-iFrame
            if (firstTimerTickFlag)
            {
                iFrame.ExecuteJavaScriptAsync("document.getElementById('account_name_text_field').focus();");
                iFrame.ExecuteJavaScriptAsync("document.getElementById('account_name_text_field').value=" + '\'' + "luvaihassanali@gmail" + '\'');
                SendKeys.SendWait(".com");
                SendKeys.SendWait("{ENTER}");
                setupTimer.Interval = 1000;
                firstTimerTickFlag = false;
            }
            else
            {
                string secret = File.ReadAllText("secret.txt");
                //To-do: wrap in await
                iFrame.ExecuteJavaScriptAsync("document.getElementById('password_text_field').focus();");
                iFrame.ExecuteJavaScriptAsync("document.getElementById('password_text_field').value=" + '\'' + secret + '\'');
                SendKeys.SendWait("0");
                SendKeys.SendWait("{ENTER}");
                setupTimer.Stop();
            }
        }

        private void InitializeWord(bool import)
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document currDoc = wordApp.Documents.Add();
            //object missing = System.Reflection.Missing.Value;

            // Word.Paragraph para1 = currDoc.Content.Paragraphs.Add(ref missing);

            /*para1.KeepTogether = -1;
            para1.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            para1.Format.SpaceBefore = 0.0f;
            para1.Format.SpaceAfter = 0.0f;
            para1.Range.Font.Name = "Arial";
            para1.Range.Font.Size = 12;
            para1.Range.Paste();*/

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

            if (import)
            {
                richTextBox1.Invoke(new MethodInvoker(delegate
                {
                    richTextBox1.Focus();
                    richTextBox1.SelectAll();
                    richTextBox1.Paste();
                }));

                Console.WriteLine("Word quit");
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
                //firstTimerTickFlag = true;
                //onNotesPage = false;
                InitializeBrowser();
            }

            if (import)
            {
                richTextBox1.Invoke(new MethodInvoker(delegate
                {
                    FixSpacing(richTextBox1, ">", "→");
                }));
            } 
            else
            {
                importing = false;
                if(this.InvokeRequired)
                {
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
                        browser.GetFocusedFrame().Paste();
                    }));
                } else
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
                    browser.GetFocusedFrame().Paste();
                }
            }
            testing here
            // make shit await
            // 
            //saving -> copy from app into word then word into notes
            //find a way to fix spacing... once pasted edit header?
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
            //Replace all → with tab + →

            int subHeaderStart = rtb.SelectionStart, subHeaderStartIndex = 0, subHeaderIndex;
            while ((subHeaderIndex = rtb.Text.IndexOf(subTarget, subHeaderStartIndex)) != -1)
            {
                rtb.SelectionStart = subHeaderIndex;
                rtb.SelectionLength = 1;
                rtb.SelectedText = "	*";
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
                rtb.SelectedText = "→";
                headerStartIndex = headerIndex + target.Length;
            }
        }


        //OnKeyEvent: KeyType: RawKeyDown 0x11 Modifiers: ControlDown, IsLeft
        //OnKeyEvent: KeyType: KeyUp 0x43 Modifiers: ControlDown
        //OnKeyEvent: KeyType: KeyUp 0x11 Modifiers: IsLeft
        private void SendCopyKeys(IBrowser iBrowser)
        {
            KeyEvent k = new KeyEvent
            {
                WindowsKeyCode = 0x11,
                Modifiers = CefEventFlags.ControlDown | CefEventFlags.IsLeft,
                //FocusOnEditableField = true,
                IsSystemKey = false,
                Type = KeyEventType.RawKeyDown
            };

            iBrowser.GetHost().SendKeyEvent(k);

            //Thread.Sleep(100);

            k = new KeyEvent
            {
                WindowsKeyCode = 0x43,
                Modifiers = CefEventFlags.ControlDown,
                //FocusOnEditableField = true,
                IsSystemKey = false,
                Type = KeyEventType.KeyUp
            };

            iBrowser.GetHost().SendKeyEvent(k);

            //Thread.Sleep(100);

            k = new KeyEvent
            {
                WindowsKeyCode = 0x11,
                Modifiers = CefEventFlags.IsLeft,
                //FocusOnEditableField = true,
                IsSystemKey = false,
                Type = KeyEventType.KeyUp
            };

            iBrowser.GetHost().SendKeyEvent(k);
        }

        private void ParseSource()
        {
            //Console.WriteLine(pageSource);
            //<header contenteditable="false">⇨</header>
            //<footer contenteditable="false">⇨</footer>
            string[] temp = pageSource.Split(new[] { "<header contenteditable=\"false\">⇨</header>" }, StringSplitOptions.None);
            string buffer = temp[1];
            temp = buffer.Split(new[] { "<footer contenteditable=\"false\">⇦</footer>" }, StringSplitOptions.None);
            buffer = temp[0];
            Console.WriteLine(buffer);
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
                //richTextBox1.Focus();
                //richTextBox1.SelectionStart = richTextBox1.Text.Length;
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

        private void OnSave(object sender, EventArgs e)
        {
            richTextBox1.Focus();
            richTextBox1.SelectAll();
            richTextBox1.Copy();

            InitializeWord(false);

            //InitializeBrowser();
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
            //Console.WriteLine("OnKeyEvent: KeyType: {0} 0x{1:X} Modifiers: {2}", type, windowsKeyCode, modifiers);
            return false;
        }

        public bool OnPreKeyEvent(IWebBrowser browserControl, IBrowser browser, KeyType type, int windowsKeyCode, int nativeKeyCode, CefEventFlags modifiers, bool isSystemKey, ref bool isKeyboardShortcut)
        {
            /*if (windowsKeyCode == 48 && modifiers == CefEventFlags.ControlDown)
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
            }*/
            return false;
        }
    }
}

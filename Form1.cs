using System;
using System.Drawing;
using System.Windows.Forms;
using ListApp.Properties;
using CefSharp.WinForms;
using CefSharp;
using System.Timers;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Resources;
using System.Configuration;

namespace ListApp
{
    public partial class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int gripOffset = 16;
        private const int menuBarOffset = 32;

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private MenuItem exitMenuItem;
        private MenuItem shutdownMenuItem;
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
        private Font contextMenuFont = new Font("Arial", 12F, FontStyle.Regular);

        public Form1()
        {
            InitializeComponent();
            InitializeBrowser();
            InitializeGui();
        }

        #region General form functions

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!shutdown)
            {
                e.Cancel = true;
                Save();
                System.Diagnostics.Debug.WriteLine("saving");
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
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            mouse_event(0x02 | 0x04, X, Y, 0, 0); // MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP
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
                    System.Diagnostics.Debug.WriteLine("frameLoadEnd importing");
                    setupTimer.Start();

                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("frameLoadEnd exporting");
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
                    System.Diagnostics.Debug.WriteLine("FrameLoadEnd main");
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

                System.Diagnostics.Debug.WriteLine("Word quit export");
                wordApp.Quit(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                setupTimer.Stop();
                setupTimer.AutoReset = false;

                this.Invoke(new MethodInvoker(delegate
                {
                    browser.Dispose();
                    System.Diagnostics.Debug.WriteLine("CEF shutdown");
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

                    //To-do: Highlight?
                    for (int i = 1; i < currDoc.Characters.Count; i++)
                    {
                        Word.Range currCharRange = currDoc.Characters[i];
                        string currCharStr = currCharRange.Text;
                        char currChar = currCharStr.ToCharArray()[0];
                        int currNumRep = currChar - '0';
                        int threadSleep = Int32.Parse(ConfigurationManager.AppSettings["Sleep"]);
                        System.Diagnostics.Debug.WriteLine(i + " " + currCharStr + " " + currChar + " " + currNumRep);

                        if (currNumRep == -35)
                        {
                            //if (i == currDoc.Characters.Count - 1 || i == currDoc.Characters.Count - 2) continue;
                            System.Threading.Thread.Sleep(threadSleep);
                            SendKeys.Send("{ENTER}");
                            System.Threading.Thread.Sleep(threadSleep);
                            continue;
                        }

                        System.Threading.Thread.Sleep(threadSleep);
                        currCharRange.Copy();
                        System.Threading.Thread.Sleep(threadSleep);
                        browser.GetFocusedFrame().Paste();
                        System.Threading.Thread.Sleep(threadSleep);
                    }
                }));

                System.Diagnostics.Debug.WriteLine("end of export");
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
                    System.Diagnostics.Debug.WriteLine("import copy");
                }));

                setupTimer.Stop();

                this.Invoke(new MethodInvoker(delegate
                {
                    InitializeTextbox();
                }));

                InitializeWord(true);

                browser.Invoke(new MethodInvoker(delegate
                {
                    browser.Visible = false;
                    this.Controls.Remove(browser);
                    browser.Dispose();
                }));

                return;

                #endregion

            }
        }

        #endregion

        #region Microsoft Word

        private void InitializeWord(bool import)
        {
            wordApp = new Word.Application();
            wordApp.Visible = false;

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
            System.Diagnostics.Debug.WriteLine("currdoc copy");

            if (import)
            {
                richTextBox1.Invoke(new MethodInvoker(delegate
                {
                    richTextBox1.Focus();
                    richTextBox1.SelectAll();
                    richTextBox1.Paste();

                }));

                System.Diagnostics.Debug.WriteLine("Word quit import");
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

        #region Gui functions 

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

        private void InitializeGui()
        {
            this.Location = Settings.Default.WinLoc;
            this.Size = Settings.Default.WinSize;
            this.Opacity = Settings.Default.Opacity;

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            trayMenu = new ContextMenu();

            exitMenuItem = new MenuItem();
            exitMenuItem.Text = "  Exit Application"; 
            exitMenuItem.Click += new EventHandler(OnExit);
            exitMenuItem.OwnerDraw = true;
            exitMenuItem.DrawItem += new DrawItemEventHandler(DrawExitMenuItem);
            exitMenuItem.MeasureItem += new MeasureItemEventHandler(MeasureExitMenuItem);

            /*infoMenuItem = new MenuItem();
            infoMenuItem.Text = " Information"; // " Info"
            infoMenuItem.Click += new EventHandler(OnInfo);
            infoMenuItem.OwnerDraw = true;
            infoMenuItem.DrawItem += new DrawItemEventHandler(DrawInfoMenuItem);
            infoMenuItem.MeasureItem += new MeasureItemEventHandler(MeasureInfoMenuItem);*/

            shutdownMenuItem = new MenuItem();
            shutdownMenuItem.Text = "  System Shutdown";
            shutdownMenuItem.Click += new EventHandler(OnShutdown);
            shutdownMenuItem.OwnerDraw = true;
            shutdownMenuItem.DrawItem += new DrawItemEventHandler(DrawShutdownMenuItem);
            shutdownMenuItem.MeasureItem += new MeasureItemEventHandler(MeasureShutdownMenuItem);

            trayMenu.MenuItems.AddRange(new MenuItem[]
            {
                exitMenuItem, shutdownMenuItem //infoMenuItem, 
            });

            trayIcon = new NotifyIcon();
            trayIcon.Text = "Notepad";
            trayIcon.Icon = new Icon("notepad.ico");
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            trayIcon.MouseClick += new MouseEventHandler(trayIcon_Click);

            contextMenu = new ContextMenuStrip();
            contextMenu.BackColor = SystemColors.Menu; //Color.FromArgb(242, 242, 242);

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

        #region Context menu draw/measure functions

        // https://stackoverflow.com/questions/6623672/how-to-put-an-icon-in-a-menuitem
        // https://www.codeproject.com/Articles/4332/Putting-Images-Next-To-MenuItems-In-A-Menu-in-C#_articleTop

        #region shutdown 

        private void MeasureShutdownMenuItem(object sender, MeasureItemEventArgs e)
        {
            MenuItem shutdownMenuItem = (MenuItem)sender;
            // Get standard menu font so that the text in this
            // menu rectangle doesn't look funny with a
            // different font
            Font menuFont = contextMenuFont;

            StringFormat stringFormat = new StringFormat();
            SizeF sizeFloat = e.Graphics.MeasureString(shutdownMenuItem.Text, menuFont, 1000, stringFormat);

            // Get image so size can be computed
            Bitmap bitmapImage = Properties.Resources.power_grey;

            // Add image height and width  to the text height and width when 
            // drawn with selected font (got that from measurestring method)
            // to compute the total height and width needed for the rectangle
            e.ItemWidth = (int)Math.Ceiling(sizeFloat.Width) + bitmapImage.Width; 
            e.ItemHeight = bitmapImage.Height; //(int)Math.Ceiling(sizeFloat.Height)
        }

        private void DrawShutdownMenuItem(object sender, DrawItemEventArgs e)
        {
            MenuItem shutdownMenuItem = (MenuItem)sender;

            // Get standard menu font so that the text in this
            // menu rectangle doesn't look funny with a
            // different font
            Font menuFont = contextMenuFont;

            // Get a brush to use for painting
            SolidBrush menuBrush = null;

            // Determine menu brush for painting
            if (shutdownMenuItem.Enabled == false)
            {
                // disabled text if menu item not enabled
                menuBrush = new SolidBrush(SystemColors.GrayText);
            }
            else // Normal (enabled) text
            {
                if ((e.State & DrawItemState.Selected) != 0)
                {
                    // Text color when selected (highlighted)
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
                else
                {
                    // Text color during normal drawing
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
            }

            // Center the text portion (out to side of image portion)
            StringFormat stringFormat = new StringFormat();
            //stringFormat.LineAlignment = System.Drawing.StringAlignment.Center;

            // Get image associated with this menu item
            Bitmap bitmapImage = Properties.Resources.power_grey;

            // Rectangle for image portion
            Rectangle rectImage = e.Bounds;

            // Set image rectangle same dimensions as image
            rectImage.Width = bitmapImage.Width;
            rectImage.Height = bitmapImage.Height;

            // Rectanble for text portion
            Rectangle rectText = e.Bounds;

            // set wideth to x value of text portion
            rectText.X += rectImage.Width;

            // Start Drawing the menu rectangle

            // Fill rectangle with proper background color
            // [use this instead of e.DrawBackground() ]
            if ((e.State & DrawItemState.Selected) != 0)
            {
                // Selected color
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(222, 222, 222)), e.Bounds);
                bitmapImage = Properties.Resources.power_red;
            }
            else
            {
                // Normal background color (when not selected)
                e.Graphics.FillRectangle(SystemBrushes.Menu, e.Bounds);
            }

            // Draw image portion
            e.Graphics.DrawImage(bitmapImage, rectImage);

            // Draw rectangle portion
            //
            // text portion
            // using menu font
            // using brush determined earlier
            // Start at offset of image rect already drawn
            // Total height,divided to be centered
            // Formated string
            e.Graphics.DrawString(shutdownMenuItem.Text,
                   menuFont,
                   menuBrush,
                   e.Bounds.Left + bitmapImage.Width,
                   e.Bounds.Top + ((e.Bounds.Height - menuFont.Height) / 2),
                   stringFormat);
        }

        #endregion

        #endregion

        #region info

        private void MeasureInfoMenuItem(object sender, MeasureItemEventArgs e)
        {
            MenuItem infoMenuItem = (MenuItem)sender;
            Font menuFont = contextMenuFont;
            StringFormat stringFormat = new StringFormat();
            SizeF sizeFloat = e.Graphics.MeasureString(infoMenuItem.Text, menuFont, 1000, stringFormat);

            // Get image so size can be computed
            Bitmap bitmapImage = Properties.Resources.copy;

            e.ItemWidth = (int)Math.Ceiling(sizeFloat.Width) + bitmapImage.Width;
            e.ItemHeight = (int)Math.Ceiling(sizeFloat.Height) + bitmapImage.Height;
        }

        private void DrawInfoMenuItem(object sender, DrawItemEventArgs e)
        {
            MenuItem infoMenuItem = (MenuItem)sender;

            // Default menu font
            Font menuFont = contextMenuFont;
            SolidBrush menuBrush = null;

            // Determine menu brush for painting
            if (infoMenuItem.Enabled == false)
            {
                // disabled text
                menuBrush = new SolidBrush(SystemColors.GrayText);
            }
            else // Normal (enabled) text
            {
                if ((e.State & DrawItemState.Selected) != 0)
                {
                    // Text color when selected (highlighted)
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
                else
                {
                    // Text color during normal drawing
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
            }

            // Center the text portion (out to side of image portion)
            StringFormat stringFormat = new StringFormat();
            stringFormat.LineAlignment = System.Drawing.StringAlignment.Center;

            // Image for this menu item
            Bitmap bitmapImage = Properties.Resources.copy;

            // Rectangle for image portion
            Rectangle rectImage = e.Bounds;

            // Set image rectangle same dimensions as image
            rectImage.Width = bitmapImage.Width;
            rectImage.Height = bitmapImage.Height;
            Rectangle rectText = e.Bounds;
            rectText.X += rectImage.Width;

            // Start Drawing the menu rectangle

            // Fill rectangle with proper background [use this instead of e.DrawBackground() ]
            if ((e.State & DrawItemState.Selected) != 0)
            {
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(222, 222, 222)), e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(SystemBrushes.Menu, e.Bounds);
            }

            e.Graphics.DrawImage(bitmapImage, rectImage);
            e.Graphics.DrawString(infoMenuItem.Text,
                menuFont, menuBrush,
                e.Bounds.Left + bitmapImage.Width,
                e.Bounds.Top + ((e.Bounds.Height - menuFont.Height) / 2),
                stringFormat);
        }

        #endregion

        #region exit

        private void MeasureExitMenuItem(object sender, MeasureItemEventArgs e)
        {
            MenuItem exitMenuItem = (MenuItem)sender;
            Font menuFont = contextMenuFont;
            StringFormat stringFormat = new StringFormat();
            SizeF sizeFloat = e.Graphics.MeasureString(exitMenuItem.Text, menuFont, 1000, stringFormat);

            // Get image so size can be computed
            Bitmap bitmapImage = Properties.Resources.close_grey;

            e.ItemWidth = (int)Math.Ceiling(sizeFloat.Width) + bitmapImage.Width;
            e.ItemHeight = bitmapImage.Height; //(int)Math.Ceiling(sizeFloat.Height) 
        }

        private void DrawExitMenuItem(object sender, DrawItemEventArgs e)
        {
            MenuItem exitMenuItem = (MenuItem)sender;

            // Default menu font
            Font menuFont = contextMenuFont;
            SolidBrush menuBrush = null;

            // Determine menu brush for painting
            if (exitMenuItem.Enabled == false)
            {
                // disabled text
                menuBrush = new SolidBrush(SystemColors.GrayText);
            }
            else // Normal (enabled) text
            {
                if ((e.State & DrawItemState.Selected) != 0)
                {
                    // Text color when selected (highlighted)
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
                else
                {
                    // Text color during normal drawing
                    menuBrush = new SolidBrush(SystemColors.MenuText);
                }
            }

            // Center the text portion (out to side of image portion)
            StringFormat stringFormat = new StringFormat();
            //stringFormat.LineAlignment = System.Drawing.StringAlignment.Center;

            // Image for this menu item
            Bitmap bitmapImage = Properties.Resources.close_grey;

            // Rectangle for image portion
            Rectangle rectImage = e.Bounds;

            // Set image rectangle same dimensions as image
            rectImage.Width = bitmapImage.Width;
            rectImage.Height = bitmapImage.Height;
            Rectangle rectText = e.Bounds;
            rectText.X += rectImage.Width;

            // Start Drawing the menu rectangle

            // Fill rectangle with proper background 
            // [use this instead of e.DrawBackground() ]
            if ((e.State & DrawItemState.Selected) != 0)
            {
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(222, 222, 222)), e.Bounds);
                bitmapImage = Properties.Resources.close_red;
            }
            else
            {
                e.Graphics.FillRectangle(SystemBrushes.Menu, e.Bounds);
            }

            e.Graphics.DrawImage(bitmapImage, rectImage);
            e.Graphics.DrawString(exitMenuItem.Text,
                menuFont,
                menuBrush,
                e.Bounds.Left + bitmapImage.Width,
                e.Bounds.Top + ((e.Bounds.Height - menuFont.Height) / 2),
                stringFormat);
        }

        #endregion

        #endregion

        #endregion

        #region Rich text box functions

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

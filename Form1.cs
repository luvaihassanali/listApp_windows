using System;
using System.Drawing;
using System.Windows.Forms;
using ListApp.Properties;

namespace ListApp
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private Boolean isVisible = true;
        public Form1()
        {
            InitializeComponent();
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Notepad";
            trayIcon.Icon = new Icon("notepad.ico");
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            trayIcon.Click += new EventHandler(trayIcon_Click);
            this.richTextBox1.LoadFile("D:\\notes.txt", RichTextBoxStreamType.PlainText);
            this.Location = Settings.Default.WinLoc;
            this.Size = Settings.Default.WinSize;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            e.Cancel = true;
            Visible = false;
            ShowInTaskbar = false;
        }

        private void trayIcon_Click(object sender, System.EventArgs e)
        {
            if (isVisible)
            {
                Visible = false;
                ShowInTaskbar = false;
                isVisible = false;
                return;
            }
            else
            {
                Visible = true;
                ShowInTaskbar = false;
                isVisible = true;
                this.WindowState = FormWindowState.Minimized;
                this.Show();
                this.WindowState = FormWindowState.Normal;

            }
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window. 
            ShowInTaskbar = false; // Remove from taskbar.
            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            this.richTextBox1.SaveFile("D:/notes.txt", RichTextBoxStreamType.PlainText);
            Settings.Default.WinLoc = this.Location;
            Settings.Default.WinSize = this.Size;
            Settings.Default.Save();
            Application.Exit();
            System.Environment.Exit(1);
        }

        private static int WM_QUERYENDSESSION = 0x11;
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            if (m.Msg == WM_QUERYENDSESSION)
            {
                this.richTextBox1.SaveFile("D:/notes.txt", RichTextBoxStreamType.PlainText);
                Settings.Default.WinLoc = this.Location;
                Settings.Default.WinSize = this.Size;
                Settings.Default.Save();
                Application.Exit();
                System.Environment.Exit(1);
            }
            base.WndProc(ref m); // If this is WM_QUERYENDSESSION, the closing event should be raised in the base WndProc.  
        }   
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                //e.Cancel = true;
                Visible = false;
                ShowInTaskbar = false;
                isVisible = false;
                return;
            }
            else
            {
                Visible = true;
                ShowInTaskbar = true;
                isVisible = true;
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
            Application.Exit();
        }
    }
}

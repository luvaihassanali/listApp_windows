using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using ListApp.Properties;

namespace ListApp
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
        private Boolean isVisible = true;
        private const int gripOffset = 16;   
        private const int yOffset = 32;
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
            trayIcon.Click += new EventHandler(trayIcon_Click);
            
            if(!File.Exists("notes.rtf"))
            {
                FileStream fs = System.IO.File.Create("notes.rtf");
                fs.Close();
            } 
            else
            {
                this.richTextBox1.LoadFile("notes.rtf", RichTextBoxStreamType.RichText);
            }

            this.Location = Settings.Default.WinLoc;
            this.Size = Settings.Default.WinSize;
            this.Opacity = Settings.Default.Opacity;

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);
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
            Visible = false;  
            ShowInTaskbar = false;
            base.OnLoad(e);
        }


        private void Form1_Deactivate(object sender, EventArgs e) {
            this.richTextBox1.SaveFile("notes.rtf", RichTextBoxStreamType.RichText);
            Settings.Default.WinLoc = this.Location;
            Settings.Default.WinSize = this.Size;
            Settings.Default.Opacity = this.Opacity;
            Settings.Default.Save();

            //Being extra... make sure Memory usage stays below 6mb in task mgr... 
            //Without these lines Memory is capped at 10mb anyways...
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void OnInfo(object sender, EventArgs e)
        {
            MessageBox.Show("Ctrl + b: Bold/Unbold \nCtrl + s: Opacity down\nCtrl + d: Opacity up");
        }

        private void OnExit(object sender, EventArgs e)
        {
            this.richTextBox1.SaveFile("notes.rtf", RichTextBoxStreamType.RichText);

            Settings.Default.WinLoc = this.Location;
            Settings.Default.WinSize = this.Size;
            Settings.Default.Opacity = this.Opacity;

            trayIcon.Visible = false;
            trayIcon.Dispose();

            Settings.Default.Save();
            Application.Exit();
            System.Environment.Exit(1);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rc = new Rectangle(this.ClientSize.Width - gripOffset, this.ClientSize.Height - gripOffset, gripOffset, gripOffset);
            ControlPaint.DrawSizeGrip(e.Graphics, this.BackColor, rc);
            //rc = new Rectangle(0, 0, this.ClientSize.Width, cCaption);
            //e.Graphics.FillRectangle(Brushes.DarkBlue, rc);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84)
            {  // Trap WM_NCHITTEST
                Point pos = new Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);
                if (pos.Y < yOffset)
                {
                    m.Result = (IntPtr)2;  // HTCAPTION
                    return;
                }
                if (pos.X >= this.ClientSize.Width - gripOffset && pos.Y >= this.ClientSize.Height - gripOffset)
                {
                    m.Result = (IntPtr)17; // HTBOTTOMRIGHT
                    return;
                }
            }  
            base.WndProc(ref m);
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            Cursor.Current = Cursors.No;
        }

        private void richTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Console.WriteLine(e.KeyChar);
            if(e.KeyChar == '')
            {
                this.Opacity -= 0.05;
            } 
            if(e.KeyChar == '')
            {
                this.Opacity += 0.05;
            }
            if(e.KeyChar == '')
            {
                if(richTextBox1.SelectionFont.Bold)
                {
                    richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Regular);
                } 
                else
                {
                    richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Bold);
                }                
            }
            if(e.KeyChar == '')
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
        }
    }
}

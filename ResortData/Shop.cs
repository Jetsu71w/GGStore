using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ResortData
{
    public partial class Shop : Form
    {
        public Shop()
        {
            InitializeComponent();
        }
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
                 
        }

        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void bunifuThinButton22_Click(object sender, EventArgs e)
        {
            st_mouse s = new st_mouse();
            s.TopLevel = false;
            subpanel1.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            st_keyboard s = new st_keyboard();
            s.TopLevel = false;
            subpanel1.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            st_headphone s = new st_headphone();
            s.TopLevel = false;
            subpanel1.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bunifuThinButton25_Click(object sender, EventArgs e)
        {
            st_microphone s = new st_microphone();
            s.TopLevel = false;
            subpanel1.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu f2 = new Menu();
            f2.Show();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
            {
                ReleaseCapture();
                SendMessage(this.Handle, 0xA1, 0x2, 0);
            }
        }
    }
    }

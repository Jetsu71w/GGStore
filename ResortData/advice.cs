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
    public partial class advice : Form
    {
        public advice()
        {
            InitializeComponent();
        }
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void วิธีการสั่งซื้อ_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu f2 = new Menu();
            f2.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bunifuThinButton22_Click(object sender, EventArgs e)
        {
            advice1 s = new advice1();
            s.TopLevel = false;
            panel3.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            advice1 s = new advice1();
            s.TopLevel = false;
            panel3.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton26_Click(object sender, EventArgs e)
        {
            advice2 s = new advice2();
            s.TopLevel = false;
            panel3.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton25_Click(object sender, EventArgs e)
        {
            advice3 s = new advice3();
            s.TopLevel = false;
            panel3.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            advice4 s = new advice4();
            s.TopLevel = false;
            panel3.Controls.Add(s);
            s.BringToFront();
            s.Show();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

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

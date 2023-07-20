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
using MySql.Data.MySqlClient;

namespace ResortData
{
    public partial class Insert : Form
    {
        public Insert()
        {
            InitializeComponent();
        }
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
            {
                ReleaseCapture();
                SendMessage(this.Handle, 0xA1, 0x2, 0);
            }
        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            MySqlConnection conn;
            string server = "localhost";
            string database = "listorder";
            string uid = "root";
            string password = "12345678";
            string connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";
            conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();
                MessageBox.Show("Connect Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connect Fail");
            }
            //---------------------------------------------
            try
            {
                string sqlCmd = "INSERT INTO item01 (categories,Itemname,price)" + "VALUES(@p1,@p2,@p3)";
                MySqlCommand cmd = new MySqlCommand(sqlCmd, conn);
                cmd.Parameters.AddWithValue("@p1", textBox1.Text);
                cmd.Parameters.AddWithValue("@p2", textBox2.Text);
                cmd.Parameters.AddWithValue("@p3", textBox3.Text);

                cmd.ExecuteNonQuery();
                MessageBox.Show("เพิ่มข้อมูลสำเร็จ");


                conn.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            this.Hide();
            exorders f7 = new exorders();
            f7.Show();
        }
    }
}

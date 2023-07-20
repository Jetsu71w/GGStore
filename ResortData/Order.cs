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
using Excel = Microsoft.Office.Interop.Excel;

namespace ResortData
{
    public partial class Order : Form
    {
        public Order()
        {
            InitializeComponent();
        }
        public string txt;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void หมวดหมู่สินค้า_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu f2 = new Menu();
            f2.Show();
        }

        private void bunifuThinButton22_Click(object sender, EventArgs e)
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
                string sqlCmd = "INSERT INTO orders ( Itemname,Name, Surname,Address,province,zipcode,tel)" + "VALUES (@p1, @p2, @p3,@p4,@p5,@p6,@p7)";
                MySqlCommand cmd = new MySqlCommand(sqlCmd, conn);
                cmd.Parameters.AddWithValue("@p1", label7.Text);
                cmd.Parameters.AddWithValue("@p2", textBox1.Text);
                cmd.Parameters.AddWithValue("@p3", textBox2.Text);
                cmd.Parameters.AddWithValue("@p4", textBox3.Text);
                cmd.Parameters.AddWithValue("@p5", textBox4.Text);
                cmd.Parameters.AddWithValue("@p6", textBox5.Text);
                cmd.Parameters.AddWithValue("@p7", textBox6.Text);
                cmd.ExecuteNonQuery();

                conn.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            this.Close();
        }

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

        private void label7_Click(object sender, EventArgs e)
        {


        }

        private void label7_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click_2(object sender, EventArgs e)
        {

        }

        private void Order_Load(object sender, EventArgs e)
        {
            label7.Text = txt;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown1.Value < 1) numericUpDown1.Value = 1;
            cal();
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            //เมาส์
            if (comboBox2.SelectedIndex == 0)
            {

                comboBox1.Items.Add("MOUSE CORSAIR SABRE RGB PRO CHAMPION SERIES 1,990.00");
                comboBox1.Items.Add("MOUSE EVGA X17 GAMING GREY 2,990.00");
                comboBox1.Items.Add("MOUSE LOGITECH G304 LIGHTSPEED WIRELESS KDA LIMITED EDITION 1,590.00 บาท");
                comboBox1.Items.Add("MOUSE NUBWO PLESIOS NM-89M PINK 390.00 บาท");
                comboBox1.Items.Add("MOUSE RAZER DEATHADDER ELITE 1,990.00");
                comboBox1.SelectedIndex = 0;
            }
            //คีย์บอร์ด
            if (comboBox2.SelectedIndex == 1)
            {
                comboBox1.Items.Add("KEYBOARD AOC GK200 1,290.00 บาท");
                comboBox1.Items.Add("KEYBOARD LOGITECH G213 PRODIGY MEMBRANE (MEMBRANE) (RGB LED) (ENTH) 1,290.00 บาท");
                comboBox1.Items.Add("KEYBOARD LOGITECH G413 CARBON GAMING MECHANICAL ROMER-G TACTILE 2,190.00 บาท");
                comboBox1.Items.Add("KEYBOARD LOGITECH G512 CARBON [GX BLUE CLICKY SWITCH] THEN 2,990.00 บาท");
                comboBox1.Items.Add("KEYBOARD MSI VIGOR GK60 MECHANICAL CHERRY RED SWITCH 3,490.00 บาท");
                comboBox1.SelectedIndex = 0;
            }
            //หูฟัง
            if (comboBox2.SelectedIndex == 2)
            {
                comboBox1.Items.Add("HEADSET ASUS ROG CETRA II CORE (IN EAR) 1,990.00 บาท");
                comboBox1.Items.Add("HEADSET LOGITECH G333 BUFFY IN EAR WHITE 1,690.00 บาท");
                comboBox1.Items.Add("HEADSET LOGITECH G335 WHITE 2,090.00 บาท");
                comboBox1.Items.Add("HEADSET LOGITECH G435 LIGHTSPEED WIRELESS -BLUE & RASPBERRY 2,290 บาท");
                comboBox1.Items.Add("HEADSET RAZER HAMMERHEAD PRO V2 1,990.00 บาท");
                comboBox1.SelectedIndex = 0;
            }
            //ไมโครโฟน
            if (comboBox2.SelectedIndex == 3)
            {
                comboBox1.Items.Add("HyperX QuadCast  4,390 บาท");
                comboBox1.Items.Add("HyperX รุ่น QuadCast S  5,990 บาท");
                comboBox1.Items.Add("HyperX รุ่น SoloCast   1,990 บาท");
                comboBox1.Items.Add("Nubwo Streaming Microphone M21   550 บาท");
                comboBox1.Items.Add("RAZER Seiren Elite   5,990 บาท");
                comboBox1.SelectedIndex = 0;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (numericUpDown1.Value < 1) numericUpDown1.Value = 1;
            cal();
        }

        private void calbtn_Click(object sender, EventArgs e)
        {
            cal();
        }
        private void cal() 
        {
            string pd = comboBox1.SelectedItem.ToString();
            double price = 0;
            double qt = 0;
            if (pd == "MOUSE CORSAIR SABRE RGB PRO CHAMPION SERIES 1,990.00") price = 1990;
            if (pd == "MOUSE EVGA X17 GAMING GREY 2,990.00") price = 2990;
            if (pd == "MOUSE LOGITECH G304 LIGHTSPEED WIRELESS KDA LIMITED EDITION 1,590.00 บาท") price = 1590;
            if (pd == "MOUSE NUBWO PLESIOS NM-89M PINK 390.00 บาท") price = 390;
            if (pd == "MOUSE RAZER DEATHADDER ELITE 1,990.00") price = 1990;

            if (pd == "KEYBOARD AOC GK200 1,290.00 บาท") price = 1290;
            if (pd == "KEYBOARD LOGITECH G213 PRODIGY MEMBRANE (MEMBRANE) (RGB LED) (ENTH) 1,290.00 บาท") price = 1290;
            if (pd == "KEYBOARD LOGITECH G413 CARBON GAMING MECHANICAL ROMER-G TACTILE 2,190.00 บาท") price = 2190;
            if (pd == "KEYBOARD LOGITECH G512 CARBON [GX BLUE CLICKY SWITCH] THEN 2,990.00 บาท") price = 2990;
            if (pd == "KEYBOARD MSI VIGOR GK60 MECHANICAL CHERRY RED SWITCH 3,490.00 บาท") price = 3490;

            if (pd == "HEADSET ASUS ROG CETRA II CORE (IN EAR) 1,990.00 บาท") price = 1990;
            if (pd == "HEADSET LOGITECH G333 BUFFY IN EAR WHITE 1,690.00 บาท") price = 1690;
            if (pd == "HEADSET LOGITECH G335 WHITE 2,090.00 บาท") price = 2090;
            if (pd == "HEADSET LOGITECH G435 LIGHTSPEED WIRELESS -BLUE & RASPBERRY 2,290 บาท") price = 2290;
            if (pd == "HEADSET RAZER HAMMERHEAD PRO V2 1,990.00 บาท") price = 1990;

            if (pd == "HyperX QuadCast  4,390 บาท") price = 4390;
            if (pd == "HyperX รุ่น QuadCast S  5,990 บาท") price = 5990;
            if (pd == "HyperX รุ่น SoloCast   1,990 บาท") price = 1990;
            if (pd == "Nubwo Streaming Microphone M21   550 บาท") price = 550;
            if (pd == "RAZER Seiren Elite   5,990 บาท") price = 5990;

            qt = (double)numericUpDown1.Value;
            prc.Text = price * qt + "";
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void prc_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void bunifuThinButton24_Click_1(object sender, EventArgs e)
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
                string sqlCmd = "SELECT * FROM orders WHERE name = @name";
                MySqlCommand cmd = new MySqlCommand(sqlCmd, conn);
                cmd.Parameters.AddWithValue("@name", textBox7.Text);
                MySqlDataReader reader = cmd.ExecuteReader();

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(reader);

                dataGridView1.DataSource = dt;
                conn.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("คุณไม่ได้ลง Nuget package Microsoft Excel Interop");
                return;
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            DataTable dt = new DataTable();

            int col_i = 1;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                xlWorkSheet.Cells[1, col_i] = column.HeaderText;
                col_i++;
            }

            int row_i = 2;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                col_i = 1;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        xlWorkSheet.Cells[row_i, col_i] = cell.Value.ToString();
                    }
                    col_i++;
                }
                row_i++;
            }
            SaveFileDialog opfd = new SaveFileDialog();
            DialogResult user_choose = opfd.ShowDialog();

            if (user_choose == DialogResult.OK)
            {
                string file_name = opfd.FileName;
                xlWorkBook.SaveAs(file_name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("ไฟล์ Excel ได้ถูกสร้างขึ้นแล้ว");
        }
    }
}


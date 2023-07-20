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
    public partial class exorders : Form
    {
        public exorders()
        {
            InitializeComponent();
        }
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

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
                string sqlCmd = "SELECT * FROM Item01 WHERE categories = @categories";
                MySqlCommand cmd = new MySqlCommand(sqlCmd, conn);
                cmd.Parameters.AddWithValue("@categories", textBox1.Text);
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

        private void exorders_Load(object sender, EventArgs e)
        {

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
                
                string sqlCmd = "DELETE FROM Item01 WHERE  Itemname = '" + textBox2.Text + "' AND price = '" + textBox3.Text + "' ";
         
                MySqlCommand cmd = new MySqlCommand(sqlCmd, conn);

                cmd.Parameters.AddWithValue("Itemname", textBox2.Text);
                cmd.Parameters.AddWithValue("price", textBox3.Text);

                cmd.ExecuteNonQuery(); 
                MessageBox.Show("ลบสำเร็จ");
                conn.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            this.Hide();
            Menu f2 = new Menu();
            f2.Show();
        }

        private void bunifuThinButton25_Click(object sender, EventArgs e)
        {
            this.Hide();
            Insert f6 = new Insert();
            f6.Show();
        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

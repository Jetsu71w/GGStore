using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResortData
{
    public partial class st_headphone : Form
    {
        public st_headphone()
        {
            InitializeComponent();
        }

        private void buy1_Click(object sender, EventArgs e)
        {
            Order f4 = new Order();
            f4.txt = label1.Text;
            f4.ShowDialog();
        }

        private void buy2_Click(object sender, EventArgs e)
        {
            Order f4 = new Order();
            f4.txt = label2.Text;
            f4.ShowDialog();
        }

        private void buy3_Click(object sender, EventArgs e)
        {
            Order f4 = new Order();
            f4.txt = label3.Text;
            f4.ShowDialog();
        }

        private void buy5_Click(object sender, EventArgs e)
        {
            Order f4 = new Order();
            f4.txt = label4.Text;
            f4.ShowDialog();
        }

        private void buy4_Click(object sender, EventArgs e)
        {
            Order f4 = new Order();
            f4.txt = label5.Text;
            f4.ShowDialog();
        }
    }
}

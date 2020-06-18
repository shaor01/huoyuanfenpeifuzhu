using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace huoyuanfenpeifuzhu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string x = "1";
            if (this.textBox1.Text == x)
            {


                Form f = new Form2();
                f.ShowDialog();
                this.Dispose();
            }
            else
            {
                MessageBox.Show("验证码不正确，退出");
                this.Dispose();
            }
         }
    }
}

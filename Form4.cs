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
    public delegate void pamWeituo(string [] str);
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
           
        }
        public string[] path;
        public string[] pam;   
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
          
                
        }

        private void Form4_Load(object sender, EventArgs e)
        {
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            int num=0;
            int num1 = 0;
            OpenFileDialog f = new OpenFileDialog();
            f.Filter= "Worksheet Files (*.xls)|*.xls||";
            f.Title = "选择投放文件，可以多选哦~~~~~";
            f.Multiselect = true;
            if (f.ShowDialog() == DialogResult.OK)
            {
                
                path = f.FileNames;
                if (path.Length > 1)
                {
                    foreach (string path1 in path)
                    {
                        if (path1.IndexOf("投放") > 0)
                        {
                            num += 1;
                        }
                    }
                    if (path.Length != num)
                    {
                        num = 0;
                        MessageBox.Show("请选择投放文件");
                    }
                    else
                    {
                        string path2 = path[0].Substring(0, path[0].IndexOf("放"));
                        foreach (string path1 in path)
                        {
                            if (!string.Equals(path2, path1.Substring(0, path[0].IndexOf("放"))))
                            {
                                num1 += 1;
                            }
                        }
                        if (num1 > 0)
                        {
                            num1 = 0;
                            MessageBox.Show("请选择同一投放周期的文件");

                        }
                        else
                        {
                            //label2.Text = path[0];
                            ExcelTodb2 dt = new ExcelTodb2();
                            pam = dt.Excel_cyc(path[0]);
                            dt = null;
                            for (int x = 0; x < pam.Length; x++)
                            {
                                string st = string.Format("label{0}", x + 9);
                                foreach (Control cr1 in Controls)
                                {
                                    if (cr1.Name == st)
                                    {
                                        cr1.Text = pam[x];
                                    }
                                }
                                button1.Enabled = true;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请选择两个或两个以上的策略文件");
                }
            } 
        }
        private Control  controlfind(string name)
        {
            Control cr;
            foreach (Control cr1 in  Controls)
            {
                if (cr1.Name == name)
                {
                     return cr1;
                }
                else if(cr1.Controls .Count >0)
                {
                    cr = controlfind(name);
                     if (cr != null) 
                     {
                        return cr;
                     }                    
                }
            }

            return null;
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        public pamWeituo weituo2 ;
        private void button1_Click(object sender, EventArgs e)
        {
            //向父窗体传递新增周期参数
            weituo2(pam);
            // _ = new DataSet();
            //ds = ExcelTodb2.ReadExcel(label2.Text);
            //查询是否已经导入过文件周期的数据
            ExcelTodb2 tt = new ExcelTodb2();
                string sql = string.Format ("select distinct cyclname,cyclcode from db2admin.sales_cycle where cyclcode='{0}'",label9.Text );
                // DataSet ds = tt.UpToDb2(sql);
                string[] st = tt.ArrayToItem(tt.DataTableToArray(sql));


            if (st == null)
            {
                MessageBox.Show("数据库查询错误，请联系管理员");
            }
            else if (st.Length >0)
            {
                MessageBox.Show("数据库已有相关记录");
            }
            else
            {    
                foreach (string tex in path)
                {
                    tt.UpToDb2(tex);
                }
            }
            //MessageBox.Show(ds.Tables.Count.ToString ());
            //dataGridView1.DataSource =ds.Tables ["pam"];
            // dataGridView1 .
            // DataRow [] dr = ds.Tables[0].Select("F5 like '%标准%'");               

            this.Dispose();
        }

        private void numericUpDown1_ValueChanged_1(object sender, EventArgs e)
        {

        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }
    }
}
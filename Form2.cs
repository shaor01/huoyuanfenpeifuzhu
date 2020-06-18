using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IBM.Data.DB2;

namespace huoyuanfenpeifuzhu
{
    public partial class Form2 : Form
    {
        public string[][] Item1;
        public string[] Item;
        public DataTable dt;
        public Form2()
        {
            InitializeComponent();
        }
       /* public string[] getxiaxia(string[][] st1)
        {

            Datatabletoarr sp = new Datatabletoarr();
            string[] st2 = new string[st1.Length ];
            st2=sp.Toarr1(st1);
            if (st2.Length == 0 || st2 == null)
            {
                return null;
            }
            else
            {
                return st2;
            }
        }
       /* public string[][] Getdd(string sql)
        {
            DataTable dl = new DataTable();
            Getdata gt = new Getdata();
            dl = gt.Getdat_db2(sql);
            string[][] st = null;
            Datatabletoarr st1 = new Datatabletoarr();
            st = st1.Toarry(dl);
            return st;


        }*/
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {            
            string sql = "select distinct productdesc,productcode from db2admin.view_quantities where cycle='" + comboBox1.Text + "' order by productcode";
            ExcelTodb2 dt = new ExcelTodb2();
            DataTable tb = new DataTable();
            Item1 = dt.DataTableToArray(sql);
            Item = dt.ArrayToItem(Item1);
            comboBox2.DataSource = Item;           
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string sql = "select distinct cycle from db2admin.view_quantities where year>'2019年' order by cycle desc";
            ExcelTodb2 dt = new ExcelTodb2();
            string [][] item2 = dt.DataTableToArray(sql);
            comboBox1.DataSource =dt.ArrayToItem(item2);
        }

        private  string cycformat(string cyc)//加工周期
        {
            string year1, month1, week1, cycle;
            year1 = cyc.Substring(0, cyc.IndexOf("年"));
            month1 = cyc.Substring(cyc.IndexOf("年") + 1, cyc.IndexOf("月") - cyc.IndexOf("年") - 1);
            week1 = cyc.Substring(cyc.IndexOf("第") + 1, cyc.IndexOf("周") - cyc.IndexOf("第") - 1);
            if (month1.Length < 2)
            {
                month1 = "0" + month1;
            }
            cycle = year1 + month1 + week1;
            return cycle;

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int x =Item .ToList().IndexOf(comboBox2.Text);
            string product =Item1 [x][1] ;            
            string cyc = cycformat(comboBox1.Text);
            string sql = "SELECT B.DANGWEI2 档位,COALESCE(SUM(SW.SUPPLY_COUNT), 0) 投放户数,COALESCE(SUM(SW.ORDER_COUNT), 0) 订购户数,COALESCE(SUM(SW.FULL_ORDER_COUNT), 0) 订足户数," +
            "COALESCE(SUM(SW.QTY_SUPPLY), 0) 投放量,COALESCE(SUM(SW.QTY_ORDER), 0) 订购量,COALESCE(SUM(SW.QTY_NEED), 0) 需求量," +
            "CASE WHEN  COALESCE(SUM(SW.QTY_SUPPLY), 0) > 0 THEN COALESCE(SUM(SW.QTY_SUPPLY), 0)  / COALESCE(SUM(SW.SUPPLY_COUNT), 0) ELSE 0 END 投放条数," +
            "CAST(ROUND((CASE WHEN COALESCE(SUM(SW.QTY_SUPPLY), 0) > 0 THEN COALESCE(SUM(SW.QTY_ORDER), 0) * 100.00 / COALESCE(SUM(SW.QTY_SUPPLY), 0)" +
            "ELSE 0 END), 2) AS DECIMAL(18, 2)) 订足率,COALESCE(SUM(SW.ORDER_COUNT), 0) - COALESCE(SUM(SW.FULL_ORDER_COUNT), 0) 订足户差," +
            "CASE WHEN  COALESCE(SUM(SW.QTY_SUPPLY), 0) > 0 THEN COALESCE(SUM(SW.QTY_SUPPLY), 0)  / COALESCE(SUM(SW.SUPPLY_COUNT), 0) ELSE 0 END 投放调整," +
            "CAST(ROUND((CASE WHEN COALESCE(SUM(SW.QTY_NEED), 0) > 0 THEN COALESCE(SUM(SW.QTY_ORDER), 0) * 100.00 / COALESCE(SUM(SW.QTY_NEED), 0)" +
            "ELSE 0 END), 2) AS DECIMAL(18, 2)) 需求满足率 FROM DB2ADMIN.CUSTDANGWEIMAPPING B left OUTER JOIN DB2ADMIN.CUSTTYPE_ITEM_SW SW " +
            "ON SW.CUST_TYPE = B.DANGWEI2 AND SW.CUST_TYPE != 'ZZ' AND SW.TYPE_KIND = '301FD' AND SW.ITEM_ID = '" + product + "' AND SW.DATE1 = '" + cyc + "'" +
            "GROUP BY B.DANGWEI2 ORDER BY B.DANGWEI2 DESC ";  
            ExcelTodb2 gt = new ExcelTodb2();
            dt =gt.SqlToDataTbale(sql);
            
            this.dataRefresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();
            chart1.DataSource = dt;
            chart1.Series[0].XValueMember = "档位";
            chart1.Series[0].YValueMembers = "订足率";
            //chart1.Series[0].LegendText = "订足率";
            //chart1.Series[0].ToolTip  = "#VALY";

            //chart1.Series[1].YValueMembers = "需求满足率";
            //chart1.Series[0].Legend = "投放条数";
            chart1.DataBind();
           // MessageBox.Show("ok");
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                chart1.ChartAreas[0].AxisY.IsReversed = true;
            }
            else
            {
                chart1.ChartAreas[0].AxisY.IsReversed = false ;
            }
        }
        private void dataRefresh()
        {
            chart1.Series[0].Points.Clear();
            chart1.DataSource = dt;
            chart1.Series[0].XValueMember = "档位";
            chart1.Series[0].YValueMembers = "订足率";
            chart1.DataBind();
            dataRefresh_lable ();
        }
        private void dataRefresh_lable()
        {
            label_xiaoliang.Text = dt.Compute("sum(订购量)", "true").ToString ();
            label_toufangliang.Text = dt.Compute("sum(投放量)", "true").ToString();
            label_hushu.Text = dt.Compute("sum(投放户数)", "true").ToString();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}

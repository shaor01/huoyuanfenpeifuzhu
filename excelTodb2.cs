using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IBM.Data.DB2;
using System.Data;
using System.Windows.Forms;
using System.Reflection;
using System.Data.OleDb;
using excel= Microsoft.Office.Interop.Excel;


namespace huoyuanfenpeifuzhu
{
    internal class ExcelTodb2
    {
        //public static  string cycle ;
        // public static DateTime  begindate;
        //public static DateTime enddate;
        public static DataSet ds;
        public static string[] pam;
        private static void DataSetValue(DataSet st)
        {
            ds = st;
        }
        //若传入的是excel文件目录则插入数据库，否则按照传入的语句返回查询结果
        internal   DataSet UpToDb2(string path)
        {   
            DataSet dtst = new DataSet();
            DB2Connection conn = new DB2Connection()
            {
                ConnectionString = "Database=cig;Password=Shaorong01;User ID=db2admin"
            };
            DB2Command comm = new DB2Command
            {
                Connection = conn
            };
            DB2DataAdapter codata = new DB2DataAdapter
            {
                SelectCommand = comm
            };
            DB2CommandBuilder cb = new DB2CommandBuilder(codata);//构建数据库插入语句
            try
            {
                if (path.IndexOf("xls") > 0)
                {
                    //从excel读取策略日期参数;
                    pam = Excel_cyc(path);
                    //将投放策略日期存入数据库
                    if (pam.Length == 7)
                    {
                        string selectSql = string.Format("select * from DB2ADMIN.SALES_CYCLE where cyclcode='{0}'", pam[0]);
                        DataTable temp1 = new DataTable() ;
                        comm.CommandText = selectSql;
                        codata.Fill(temp1);
                        if (temp1.Rows.Count < 1)
                        {
                            string insertSql = string.Format("INSERT INTO DB2ADMIN.SALES_CYCLE(YEAR, MONTH, WEEK, BEGINDATE, ENDDATE, " +
                            "CYCLCODE, CYCLNAME) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", pam[1], pam[2], pam[3], pam[4], pam[5], pam[0], pam[6]);
                            comm.CommandText = insertSql;
                            conn.Open();
                            comm.ExecuteNonQuery();
                            conn.Close();
                        }
                        temp1.Dispose();
                    }
                    string cycle = pam[0];
                    comm.CommandText = "SELECT CYCLE, PRODUCT_CODE, PRODUCT_NAME, SALE_DEPT, TYPE_NAME, CUST_TYPE, SGP_NUM ,QR_ID FROM DB2ADMIN.QUANTITIES_RECORD_NEW WHERE CYCLE='" + cycle + "'";
                    codata.Fill(dtst, "excel");
                    DataSet stt1 = new DataSet();
                    //从excel读取策略内容
                    stt1 = ReadExcel(path);
                    if (stt1!=null)
                    {
                        //将excel内容转换为数据库结构
                        handleDataTable(stt1, cycle, ref dtst);
                    }
                    //丢弃读取的excel内容
                    stt1.Dispose();     
                    //向数据库写入
                    codata.Update(dtst.Tables["excel"].Select(null, null, DataViewRowState.Added));
                    dtst.Clear();
                }
                else
                {
                    comm.CommandText = path;
                    codata.Fill(dtst, "pam");
                }
                return dtst;
                // DataSetValue(dtst);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                codata.Dispose();
                comm.Dispose();
                conn.Dispose();
                dtst.Dispose();
            }
        }
        //按照指定的语句查询数据库，将结果输出为表格
        public  DataTable SqlToDataTbale(string sql)
        {
            if (sql != null)
            {
                _ = new DataSet();
                DataSet ss = UpToDb2(sql);
                if (ss != null)
                {
                    DataTable dt = new DataTable();
                    dt = ss.Tables["pam"];
                    ss.Dispose();
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }       
        }
        //按照指定的语句查询数据库，将结果输出为数组
        public string[][] DataTableToArray(string sql)
        {
            DataTable dt = new DataTable();
            dt = SqlToDataTbale(sql);
            if (dt == null)
            {                
                return null;
            }
            else
            {
                string[][] str = new string[dt.Rows.Count][];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    str[i] = new string[dt.Columns.Count];
                    dt.Rows[i].ItemArray.CopyTo(str[i], 0);
                }
                dt.Dispose();
                return str;
            }
        }
        //将2维数组转为1维
        public string[] ArrayToItem(string[][] str)
        {
            if (str == null)
            {
                return null;
            }
            else
            {
                string[] st = new string[str.Length];
                for (int i = 0; i < str.Length; i++)
                {
                    st[i] = str[i][0];
                }
                return st;
            }


        }

        //从excel读数据存储到数据集   
        private   DataSet  ReadExcel(string path )
        {
            DataSet ds = new DataSet();
            string constr;
            if (JongCheckExcelVer() >= 12)//"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。
            {
                constr = "provider=microsoft.ace.oledb.12.0;extended properties=excel 12.0;data source='" + path + "';";
            }
            else
            {
                constr = "provider=microsoft.jet.oledb.4.0;extended properties=excel 11.0;data source='" + path + "';";
            }

            OleDbConnection conn = new OleDbConnection
            {
                ConnectionString = constr
            };
            try
            {
                conn.Open();
                DataTable tb = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);                
                for (int x = 0; x < tb.Rows.Count ; x++)
                {
                    string sql = string.Format ( "select * from [{0}] " , tb.Rows[x][2].ToString());
                    //string sql = "select * from [" + tb.Rows[x][2].ToString() + "a3:a65530";
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, conn); 
                    //DataTable dt = new DataTable (tb.Rows[x][2].ToString());
                    ada.Fill(ds, tb.Rows[x][2].ToString()); 
                   // handleDataTable(ref dt);//处理数据表 
                    ada.Dispose();
                }
                tb.Dispose();
                return ds;
                //Console.WriteLine(tb.Rows );
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                conn.Close();
                conn.Dispose();                
                ds.Dispose();
            } 
            
        }

        //获取excel版本
        public static double JongCheckExcelVer()
        {
            Type objExcelType = Type.GetTypeFromProgID("Excel.Application");
            if (objExcelType == null)
            {
                return 0;
            }
            object objApp = Activator.CreateInstance(objExcelType);
            if (objApp == null)
            {
                return 0;
            }
            object objVer = objApp.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, objApp, null);
            double iVer = Convert.ToDouble(objVer.ToString());
            objVer = null;
            objApp = null;
            objExcelType = null;
            GC.Collect();
            return iVer;
        }

        //从excel获取的数据按照格式转换到数据库格式数据集
        private static void handleDataTable(DataSet dt,string cycle, ref DataSet rdt)
        {
            if (dt!=null)
            {
                string[] st;
                for (int x = 0; x < dt.Tables.Count; x++)
                {
                    //按照表格样式，区县+市场类型+档位的限量位置与其他不一致，因此列集合内容设置为空
                    if (dt.Tables[x].TableName.IndexOf("+") < 0)
                    {
                        st = new string[] { "档位" };
                    }
                    else if (dt.Tables[x].TableName.ToString() == "'档位+诚信等级$'")
                    {
                        st = new string[] {"A0","A1","A2","A3","A4","AA0","AA1","AA2","AA3","AA4","AAA0","AAA1","AAA2" ,"AAA3","AAA4",
                        "B0","B1","B2","B3","B4","C0","C1","C2","C3","C4","D0","D1","D2","D3","D4" };

                    }
                    else if (dt.Tables[x].TableName.ToString() == "'市场类型+档位$'")
                    {
                        st = new string[] { "城镇", "乡村" };
                    }
                    else if (dt.Tables[x].TableName.ToString() == "'区县+档位$'")
                    {
                        st = new string[] { "渑池", "陕州", "卢氏", "湖滨", "义马", "灵宝" };

                    }
                    else
                    {
                        st = null;
                    }
                    //按列循环，将列内容加入TYPENAME
                    foreach (DataRow row1 in dt.Tables[x].Rows)
                    {
                        if (!(string.IsNullOrEmpty(row1[1].ToString().Trim()) || row1[1].ToString().IndexOf("编码") > 0))
                        {
                            //此处即为区县+市场类型+档位投放量的转存
                            if (st == null)
                            {
                                string[] ddb = new string[] { "城镇", "乡村" };
                                for (int m = 0; m < 2; m++)
                                {
                                    DataRow roo = rdt.Tables[0].NewRow();
                                    roo[0] = cycle; roo[1] = row1[1];
                                    roo[2] = row1[2]; roo[3] = row1[6]; roo[4] = ddb[m]; roo[5] = row1[3]; roo[6] = row1[7 + m];
                                    rdt.Tables[0].Rows.Add(roo);
                                }
                            }
                            //此处为其他投放方式的转存
                            else
                            {
                                for (int m = 0; m < st.Length; m++)
                                {
                                    DataRow roo = rdt.Tables[0].NewRow();
                                    roo[0] = cycle; roo[1] = row1[1];
                                    roo[2] = row1[2];
                                    roo[4] = st[m];
                                    roo[5] = row1[3]; roo[6] = row1[5 + m];
                                    rdt.Tables[0].Rows.Add(roo);
                                }
                            }
                        }
                    }
                }

            }
            
        }
        private string CycleMake(string cyc)
        {
            string str1 = cyc;
            if (string.IsNullOrEmpty(str1))
            {
                 str1 = "";
            
            }
            return str1;
        }
        //获取投放周期、投放开始日期、结束日期
        public  string[] Excel_cyc(string path)
        {
            excel.Application ex = new excel.Application();
            excel._Workbook tb = ex.Workbooks.Open(path);
            excel.Worksheet wkst = tb.Worksheets[1];
            try
            {
               string cyc=wkst.Cells[2,1].value;
                
                List<string> ls = new List<string>();
                ls.Add(cyc.Substring (0,cyc.IndexOf (")")+1));
                //ls.Add(cycformat(cyc));
                ls.Add(cyc.Substring(0, cyc.IndexOf("年")+1));
                ls.Add(cyc.Substring(cyc.IndexOf("年") + 1, cyc.IndexOf("月") - cyc.IndexOf("年") ));
                ls.Add(cyc.Substring(cyc.IndexOf("第") , cyc.IndexOf("周") - cyc.IndexOf("第")+1));
                ls.Add(BeginDate(cyc).ToString ());
                ls.Add(EndDate(cyc).ToString ());
                ls.Add(cycformat(cyc));
                return ls.ToArray ();
            }
            catch (Exception dsex)
            {
                MessageBox.Show(dsex.Message);
                string[] err= { "err" };
                return err;
            }
            finally
            {
                wkst = null;
                tb.Close();
                tb = null;
                ex.Quit();
                ex = null;
            }
        }
        //加工cycle
        public  string cycformat(string cyc)//加工周期
        {
            string year1, month1, week1,cycle;
            year1 = cyc.Substring(0, cyc.IndexOf("年"));
            month1 = cyc.Substring(cyc.IndexOf("年") + 1, cyc.IndexOf("月") - cyc.IndexOf("年") - 1);
            week1 = cyc.Substring(cyc.IndexOf("第") + 1, cyc.IndexOf("周") - cyc.IndexOf("第") - 1);
            if (month1.Length < 2)
            {
                month1 = "0" + month1;
            }
            cycle = year1 + month1 + week1;
            return  cycle;

        }
        private static string BeginDate(string cyc)//加工开始日期
        {            
            string year1 = cyc.Substring(0, cyc.IndexOf("年"));
            string[] dtemp = cyc.Substring(cyc.IndexOf("(") + 1, cyc.IndexOf(")") - cyc.IndexOf("(") - 1).Split('-');
           string begindate = year1 + "-" + dtemp[0].Replace(".", "-");
            return begindate;            
        }
        private static string EndDate(string cyc)//加工结束日期
        {
            string year1 = cyc.Substring(0, cyc.IndexOf("年"));
            string[] dtemp = cyc.Substring(cyc.IndexOf("(") + 1, cyc.IndexOf(")") - cyc.IndexOf("(") - 1).Split('-');
            string  enddate = year1 + "-" + dtemp[1].Replace(".", "-");
            return enddate;
        }
      
    }
}

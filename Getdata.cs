using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IBM.Data.DB2;
using System.Data;

namespace huoyuanfenpeifuzhu
{
   internal class Getdata
    {
        public static string constrbuild(int tt)
        {
            string str = "";
            switch (tt)
            {
                case 1: 
                    str = "Database=cig;Password=Shaorong01;User ID=db2admin";
                    break;
                case 2:
                    str = "";
                    break;
                case 3:
                    str = "";
                    break;
                       
            }
            return str;
        }


        public DataTable Getdat_db2(string sql)
        {
            
            string constr = "Database=cig;Password=Shaorong01;User ID=db2admin";
            DB2Connection conn = new DB2Connection(constr);
            try
            {
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DB2Command sqlcon = new DB2Command(sql, conn);
                    DB2DataAdapter ds = new DB2DataAdapter(sqlcon);
                    DataTable dt = new DataTable();
                    ds.Fill(dt);
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                conn.Close();
            }
        }
        public DataTable Getda_acc(string sql,string path)
        {
            string constr = "Database=cig;Password=Shaorong01;User ID=db2admin";
            DB2Connection conn = new DB2Connection(constr);
            try
            {
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DB2Command sqlcon = new DB2Command(sql, conn);
                    DB2DataAdapter ds = new DB2DataAdapter(sqlcon);
                    DataTable dt = new DataTable();
                    ds.Fill(dt);
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                conn.Close();
            }

        }
        public DataTable Getda_excel(string sql, string path)
        {
            string constr = "Database=cig;Password=Shaorong01;User ID=db2admin";
            DB2Connection conn = new DB2Connection(constr);
            try
            {
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DB2Command sqlcon = new DB2Command(sql, conn);
                    DB2DataAdapter ds = new DB2DataAdapter(sqlcon);
                    DataTable dt = new DataTable();
                    ds.Fill(dt);
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public string dayTocycle(DateTime t)
        {
            string year1 = t.Year.ToString();
            string month1 = "";
            if (t.Month.ToString().Length < 2)
            {
                month1 = "0" + t.Month.ToString();
            }
            else
            {
                month1 = t.Month.ToString();
            }

            string week1 = getWeekNumInMonth(t).ToString();
            string tt = year1 + month1 + week1;
            return tt;
        }
        private  static int getWeekNumInMonth(DateTime daytime)
        {
            int dayInMonth = daytime.Day;
            //本月第一天  
            DateTime firstDay = daytime.AddDays(1 - daytime.Day);
            //本月第一天是周几  
            int weekday = (int)firstDay.DayOfWeek == 0 ? 7 : (int)firstDay.DayOfWeek;
            //本月第一周有几天  
            int firstWeekEndDay = 7 - (weekday - 1);
            //当前日期和第一周之差  
            int diffday = dayInMonth - firstWeekEndDay;
            diffday = diffday > 0 ? diffday : 1;
            //当前是第几周,如果整除7就减一天  
            int WeekNumInMonth = ((diffday % 7) == 0
                                        ? (diffday / 7 - 1)
                                        : (diffday / 7)) + 1 + (dayInMonth > firstWeekEndDay ? 1 : 0);
            return WeekNumInMonth;

        }
        public string cycformat(string cyc)
        {
            string year1, month1, week1,data1;
            year1 = cyc.Substring(0, cyc.IndexOf("年"));
            month1 = cyc.Substring(cyc.IndexOf("年") + 1, cyc.IndexOf("月") - cyc.IndexOf("年") - 1);
            week1 = cyc.Substring(cyc.IndexOf("第") + 1, cyc.IndexOf("周") - cyc.IndexOf("第") - 1);
            if (month1.Length < 2)
            {
                month1 = "0" + month1;
            }
            data1 = year1 + month1 + week1;
            return data1;
            

        }
    }




}     

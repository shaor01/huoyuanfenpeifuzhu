using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace huoyuanfenpeifuzhu
{
    class Datatabletoarr
    {
        public string[][] Toarry(System.Data.DataTable dt)
        {
            
            if (dt == null || dt.Rows.Count == 0)
            {

                return null;
            }
            else
            {
                string[][] str = new string[dt.Rows.Count][];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    str[i] = new string[dt.Columns.Count];
                    dt.Rows [i].ItemArray .CopyTo( str[i],0);
                }
                return str;
            }
        }
        public string[] Toarr1(string[][] str)
        {
            if (str.Length == 0 || str == null)
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
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenXml.ExcelHelper
{
    public class TestData
    {
        private static string _exportDir = @"E:\test";

        public static string GetNewExcelFileName(string name)
        {
            //return Path.Combine(_exportDir, DateTime.Now.ToString("yyMMdd-HHmmss") + suffix);
            return Path.Combine(_exportDir, name);
        }

        public static string GetFileName(string fileName)
        {
            return Path.Combine(_exportDir
                , fileName);
        }

        public static DataTable GetDataTable(int cols = 10, int rows = 10, string tabName = "mytable")
        {
            DataTable dt = new DataTable(tabName);
            for (int i = 0; i < cols; i++)
            {
                dt.Columns.Add("col" + i.ToString("D3"));
            }

            DataRow dr = null;
            for (int i = 0; i < rows; i++)
            {
                dr = dt.NewRow();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    dr[j] = "val-" + i + "-" + j;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        internal static DataTable NewGetDataTable(DataTable dt, string tableName)
        {
            int cols = dt.Columns.Count;
            int rows = dt.Rows.Count;
            //DataTable dt1 = dt.Copy();
            DataTable dataTable = new DataTable(tableName);


            var tableHead = dt.Columns;
            //object[] vs = new object[] { dt.Columns };
            for (int i = 0; i < cols; i++)
            {
                //dt.Columns.Remove(dt.Columns[i]);
                dataTable.Columns.Add(tableHead[i]);
            }


            return dataTable;
        }
    }
}

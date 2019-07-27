using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenXml.Sql
{
    public class SqlDA
    {
        public static SqlConnection _con;
        public static SqlCommand _cmd;

        internal static DataTable Select(string sqlstr, string sql)
        {
            try
            {
                DataTable dataTable = null;
                _con = new SqlConnection(sqlstr);
                _con.Open();
                _cmd = new SqlCommand(sql, _con);
                SqlDataAdapter sqlData = new SqlDataAdapter(_cmd);
                DataSet dataSet = new DataSet();
                sqlData.Fill(dataSet);
                if (dataSet != null)
                {
                    dataTable = dataSet.Tables[0];
                }
                _cmd.ExecuteNonQuery();
                _con.Close();
                return dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }
}

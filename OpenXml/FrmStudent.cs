using OpenXml.ExcelHelper;
using OpenXml.Sql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenXml
{
    public partial class FrmStudent : Form
    {
        private DataGridView gridView;
        static string sql = "select * from student";
        static string sqlstr = "server=LAPTOP-SM4B8RCG;database=test;user=sa;pwd=admin";

        public FrmStudent()
        {
            InitializeComponent();
            gridView = dataGridView1;
        }

        private void BtnQuery_Click(object sender, EventArgs e)
        {
            DataTable dataTable = SqlDA.Select(sqlstr,sql);
            gridView.DataSource = dataTable;
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();
            dataTable = gridView.DataSource as DataTable;
            ExprotExcel(gridView);
        }

        private void ExprotExcel(DataGridView gridView)
        {
            string tableName = "StudentInfo";

            CreatExcel.MadeTable(gridView, tableName);
        }
    }
}

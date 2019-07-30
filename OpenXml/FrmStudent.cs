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
        int buttonIndex;

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

        private void BtnOpenXmlExport_Click(object sender, EventArgs e)
        {
            buttonIndex = 0;
            DataTable dataTable = new DataTable();
            dataTable = gridView.DataSource as DataTable;
            //dataTable.TableName = "Sheet";
            //ExprotExcel(gridView);
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "导出Excel";
            saveFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx";
            DialogResult dialogResult = saveFileDialog.ShowDialog(this);
            if(dialogResult==DialogResult.OK)
            {
                CreatExcel.ExprotToExcel(dataTable, saveFileDialog.FileName,buttonIndex);
                MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK);
            }
        }

        private void BtnOpenXmlImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "导入信息表";
            openFileDialog.Filter = "Excel文档(*.xlsx)|*.xlsx";
            if(openFileDialog.ShowDialog()==DialogResult.OK)
            {
                //var dataTable = ExcelOpenXml.GetSheet(openFileDialog.FileName);
                var dataTable = ExcelOpenXml.GetSheet(openFileDialog.FileName);
                gridView.DataSource = dataTable;
            }
        }

        private void BtnNpoiExport_Click(object sender, EventArgs e)
        {
            buttonIndex = 1;
            DataTable dataTable = new DataTable();
            dataTable = gridView.DataSource as DataTable;
            dataTable.TableName = "Sheet";
            //ExprotExcel(gridView);
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "导出Excel";
            saveFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls";
            DialogResult dialogResult = saveFileDialog.ShowDialog(this);
            if (dialogResult == DialogResult.OK)
            {
                CreatExcel.ExprotToExcel(dataTable, saveFileDialog.FileName,buttonIndex);
                MessageBox.Show("导出成功", "提示", MessageBoxButtons.OK);
            }
        }

        private void BtnNpoiImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "导入信息表";
            openFileDialog.Filter = "Excel文档(*.xlsx)|*.xlsx|Excel文档(*.xls)|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //var dataTable = ExcelOpenXml.GetSheet(openFileDialog.FileName);
                var dataTable = ExcelNpoi.GetSheet(openFileDialog.FileName);
                gridView.DataSource = dataTable;
            }
        }
    }
}

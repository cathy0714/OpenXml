using OpenXml.ExcelHelper;
using OpenXml.Sql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
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
            gridView.ReadOnly = true;
            gridView.Enabled = true;
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

        private void BtnExportWord_Click(object sender, EventArgs e)
        {
            string AppPath = AppDomain.CurrentDomain.BaseDirectory;
            string templateWordPath = AppPath + "word模板.docx";
            //string dateTime = DateTime.Now.ToString("yyyy-MM-dd-HHmmss");
            string wordPath = AppPath + "word/" + DateTime.Now.ToString("yyyy-MM-dd-HHmmss") + ".docx";
            OpenXmlWordUtil.CopyFile(templateWordPath, wordPath);

            #region 插入简单文本
            // 插入简单文本
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("公司名称", "上海数慧系统技术有限公司");
            dic.Add("公司简介", "数外慧中，专精至善");
            dic.Add("路径", wordPath);
            dic.Add("导出时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            OpenXmlWordUtil.InsertText(wordPath, dic);
            #endregion

            #region 插入表格
            // 插入表格
            List<TableDataModel> tableDataModels = new List<TableDataModel>();
            for (int i=0;i<5;i++)
            {
                TableDataModel tableDataModel = new TableDataModel();
                for(int m=0;m<4;m++)
                {
                    string p = "Property" + (m + 1);
                    Type type = tableDataModel.GetType();
                    PropertyInfo propertyInfo = type.GetProperty(p);
                    propertyInfo.SetValue(tableDataModel, p + "-" + DateTime.Now.ToString("sss"), null);
                }
                tableDataModels.Add(tableDataModel);
            }

            ConfigModel configModel = new ConfigModel("table1");
            OpenXmlWordUtil.InsertTable(wordPath, configModel, tableDataModels);



            #endregion
        }
    }
}

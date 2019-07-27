using Microsoft.VisualStudio.TestTools.UnitTesting;
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
    [TestClass]
    public class CreatExcel
    {
        [TestMethod]
        public static void MadeTable(DataGridView dgv, string tableName)
        {
            var fname = TestData.GetNewExcelFileName("ExcelText.xlsx");
            var dataTable = TestData.GetDataTable(dgv, tableName);

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable.Copy());

            ExcelOpenXml.Create(fname, dataSet);
            Assert.IsTrue(File.Exists(fname));
        }
    }
}

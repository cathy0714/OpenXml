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
        public static void ExprotToExcel(DataTable dataTable, string fileName, int index)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable.Copy());
            
            //ExcelOpenXml.Create(fileName, dataSet);
            switch(index)
            {
                case 0:
                    ExcelOpenXml.Create(fileName, dataSet);
                    break;
                case 1:
                    ExcelNpoi.Create(fileName, dataSet);
                    break;
                default:
                    break;

            }
            Assert.IsTrue(File.Exists(fileName));
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml
{
    /// <summary>
    /// 常用工具类——Excel操作类
    /// </summary>
    public class NewExcelHelper
    {
        SpreadsheetDocument spreadsheetDocument;
        //WorkbookPart _workbookPart;
        //Workbook _workbook;
        string _filePath;

        public NewExcelHelper(string filePath)
        {
            _filePath = filePath;
            Open();
        }

        /// <summary>
        /// 以读写模式打开
        /// </summary>
        private void Open()
        {
            spreadsheetDocument = SpreadsheetDocument.Open(_filePath, true);
        }
    }
}

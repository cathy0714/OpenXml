using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
         * excel 对象结构
         * SpreadsheetDocument
         *   》WorkbookPart
         *       》WorksheetPart
         *           》Worksheet
         *            》SheetData
         *       》WorksheetPart
         *          》Worksheet
         *                》SheetData1
         *       》Workbook
         *           》Sheets
         *                》Sheet
         */

namespace ExcelTest
{
    /// <summary>
    /// 常用工具类——Excel操作类
    /// </summary>
    public class ExcelHelper
    {

        private SpreadsheetDocument spreadsheetDocument;
        //WorkbookPart _workbookPart;
        //Workbook _workbook;
        string _file;

        #region 全局属性及方法
        public static bool Ok = true;
        public static bool Err = false;
        public static void ShowMessage(string strMsg)
        {
            Console.WriteLine(strMsg);
        }
        #endregion

        public ExcelHelper(string file)
        {
            _file = file;
        }

        /// <summary>
        /// 以读写模式打开电子表格
        /// </summary>
        public bool Open(string filepath)
        {
            try
            {
                spreadsheetDocument = SpreadsheetDocument.Open(filepath, true);
                if (spreadsheetDocument == null)
                    return Err;
                return Ok;
            }
            catch(Exception ex)
            {
                ShowMessage("[Func:Open],Exception:" + ex.Message);
                return Err;
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <returns></returns>
        public bool Save()
        {
            try
            {
                spreadsheetDocument.WorkbookPart.Workbook.Save();
                return Ok;
            }
            catch(Exception ex)
            {
                ShowMessage("[Func:Save],Exception:" + ex.Message);
                return Err;
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <returns></returns>
        public bool Close()
        {
            try
            {
                // 先保存
                Save();
                spreadsheetDocument.Close();
                return Ok;
            }
            catch(Exception ex)
            {
                ShowMessage("[Func:Close],Exception:" + ex.Message);
                return Err;
            }
        }

        /// <summary>
        /// 根据文件路径创建电子表格
        /// </summary>
        /// <param name="filepath">路径</param>
        /// <returns></returns>
        public bool CreatExcel(string filepath)
        {
            try
            {
                spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
                // 新建工作簿部件并添加到document中
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // 在工作簿部件中加入一个工作表部件
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // 在工作簿下新建表结构
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // 创建sheet表
                Sheet sheet = new Sheet
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = UInt32Value.FromUInt32(1),
                    Name = "Sheet"
                };
                sheets.Append(sheet);// 将sheet加入到sheets

                //worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
                spreadsheetDocument.Close();

                return Ok;
            }
            catch(Exception ex)
            {
                ShowMessage("[Func:CreatExcel],Exception:" + ex.Message);
                return Err;
            }
        }
    }
}

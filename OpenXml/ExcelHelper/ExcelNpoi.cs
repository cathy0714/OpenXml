﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenXml.ExcelHelper
{
    public class ExcelNpoi
    {
        public static void Create(string filename, DataSet ds)
        {
            using (FileStream fs =
              new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = null; //创建Workbook对象
                if (Regex.IsMatch(filename, ".xlsx"))
                    workbook = new XSSFWorkbook(); //2007 excel（.xlsx）
                else
                    workbook = new HSSFWorkbook(); //2003 excel (.xls)


                for (int s = 0; s < ds.Tables.Count; s++)
                {
                    DataTable dt = ds.Tables[s];
                    var tname = dt.TableName;
                    ISheet sheet = workbook.CreateSheet(tname); //创建工作表
                    //IRow row = sheet.CreateRow(0); //在工作表中添加一行
                    //ICell cell = row.CreateCell(0); //在行中添加一列
                    //cell.SetCellValue("test"); //设置列的内容
                    int index = 0;
                    IRow row = null;
                    ICell cell = null;
                    object val = null;
                    //标题行
                    row = sheet.CreateRow(index++);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        cell = row.CreateCell(i);
                        cell.SetCellValue(dt.Columns[i].ColumnName);
                    }
                    //内容行
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        row = sheet.CreateRow(index++);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            cell = row.CreateCell(j);
                            val = dt.Rows[i][j];
                            cell.SetCellValue(val.ToString().Trim(new char[] { '0', ':' }));

                        }
                    }
                }
                workbook.Write(fs);
            }
        }


        public static DataTable GetSheet(string filename)
        {
            DataTable dt = new DataTable();


            using (FileStream fs =
              new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = null; //从流内容创建Workbook对象
                if (Regex.IsMatch(filename, ".xlsx"))
                    workbook = new XSSFWorkbook(fs); //2007 excel（.xlsx）
                else
                    workbook = new HSSFWorkbook(fs); //2003 excel (.xls)


                //ISheet sheet = workbook.GetSheetAt(0); //获取第一个工作表
                //IRow row = sheet.GetRow(0); //获取工作表第一行
                //ICell cell = row.GetCell(0); //获取行的第一列
                //string value = cell.ToString(); //获取列的值
                ISheet sheet = null;
                IRow row = null;
                ICell cell = null;
                string sheetName = workbook.GetSheetName(0);
                sheet = workbook.GetSheet(sheetName);
                int index = 0;
                int rows = sheet.LastRowNum;
                //首行标题
                for (; index < rows;)
                {
                    row = sheet.GetRow(index);
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        cell = row.Cells[i];
                        dt.Columns.Add(cell.StringCellValue);
                    }
                    index++;
                    break;
                }
                //内容行
                DataRow dr = null;
                for (int num=0; num < rows; num++)
                {
                    dr = dt.NewRow();
                    row = sheet.GetRow(index++);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        cell = row.Cells[i];
                        dr[i] = cell.ToString();
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
    }
}

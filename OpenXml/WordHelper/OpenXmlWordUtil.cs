using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml
{
    public class OpenXmlWordUtil
    {
        /// <summary>
        /// 复制模板到目标文件夹下
        /// </summary>
        /// <param name="templateWordPath"></param>
        /// <param name="wordPath"></param>
        public static void CopyFile(string templateWordPath, string wordPath)
        {
            if(File.Exists(templateWordPath))
            {
                File.Copy(templateWordPath, wordPath, true);
            }
        }

        /// <summary>
        /// 插入文本
        /// </summary>
        /// <param name="wordPath"></param>
        /// <param name="dic"></param>
        public static void InsertText(string wordPath, Dictionary<string, string> dic)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(wordPath, true))
            {
                List<BookmarkStart> allBookmarkStart = wordprocessingDocument.MainDocumentPart.RootElement.Descendants<BookmarkStart>().ToList();
                foreach(var dictionary in dic )
                {
                    foreach(var bookmarkStart in allBookmarkStart)
                    {
                        if(bookmarkStart.Name.Value==dictionary.Key)
                        {
                            InsertIntoBookmark(bookmarkStart, dictionary.Value);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 在指定书签处插入文本
        /// </summary>
        /// <param name="bookmarkStart"></param>
        /// <param name="value"></param>
        private static void InsertIntoBookmark(BookmarkStart bookmarkStart, string value)
        {
            OpenXmlElement xmlElement = bookmarkStart.NextSibling();
            while(xmlElement!=null&&!(xmlElement is BookmarkEnd))
            {
                OpenXmlElement element = xmlElement.NextSibling();
                xmlElement.Remove();
                xmlElement = element;
            }
            bookmarkStart.Parent.InsertAfter<DocumentFormat.OpenXml.Wordprocessing.Run>
                (new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(value)), bookmarkStart);
        }

        /// <summary>
        /// 插入表格数据
        /// </summary>
        /// <param name="wordPath"></param>
        /// <param name="configModel"></param>
        /// <param name="tableDataModels"></param>
        public static void InsertTable(string wordPath, ConfigModel configModel, List<TableDataModel> tableDataModels)
        {
            using (WordprocessingDocument wordprocessing = WordprocessingDocument.Open(wordPath, true))
            {
                Body body = wordprocessing.MainDocumentPart.Document.Body;
                List<BookmarkStart> bookmarkStarts = wordprocessing.MainDocumentPart.RootElement.Descendants<BookmarkStart>().ToList();

                BookmarkStart bookmarkStart = bookmarkStarts.Find(q => q.Name.Value == configModel.tableBookMark);
                if(bookmarkStart==null)
                {
                    return;
                }
                var table = bookmarkStart.Parent.Parent.Parent.Parent;
                foreach(var tablemodel in tableDataModels)
                {
                    var row = table.Elements<TableRow>().ElementAt(1).Clone() as TableRow;
                    var cells = row.Elements<TableCell>();
                    for(int i=0;i<cells.Count();i++)
                    {
                        var cell = cells.ElementAt(i);
                        var tmpPa = cell.Elements<Paragraph>().First();
                        var tmpRuns = tmpPa.Elements<Run>();

                        if (tmpRuns.Count() <= 0)
                        {
                            tmpPa.Remove();
                            tmpPa = new Paragraph(new Run(new Text(" ")));
                            //cell.RemoveAllChildren();
                            cell.Append(tmpPa);

                        }
                        var tmpRun = tmpPa.Elements<Run>().First();
                        var tmpText = tmpRun.Elements<Text>().First();

                        //获取属性值
                        Type type = tablemodel.GetType();
                        string propertyKey = "Property" + (i + 1);
                        System.Reflection.PropertyInfo propertyInfo = type.GetProperty(propertyKey); //获取指定名称的属性
                        object objValue = propertyInfo.GetValue(tablemodel, null);
                        if (objValue != null)
                        {
                            tmpText.Text = objValue.ToString();
                        }
                        else
                        {
                            tmpText.Text = "-";
                        }
                    }

                    var lastRow = table.Elements<TableRow>().Last();
                    table.InsertAfter(row, lastRow);
                }
                //最后删除startIndex行
                table.Elements<TableRow>().ElementAt(1).Remove();
            }
        }

        /// <summary>
        /// 想表格中插入数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="wordPath"></param>
        /// <param name="configModel1"></param>
        /// <param name="studentInfoModelList"></param>
        /// <param name="obj"></param>
        internal static void InsertTable<T>(string wordPath, ConfigModel configModel1, List<T> studentInfoModelList,string[] obj)
        {
            using (WordprocessingDocument wordprocessing = WordprocessingDocument.Open(wordPath, true))
            {
                Body body = wordprocessing.MainDocumentPart.Document.Body;
                List<BookmarkStart> bookmarkStarts = wordprocessing.MainDocumentPart.RootElement.Descendants<BookmarkStart>().ToList();

                BookmarkStart bookmarkStart = bookmarkStarts.Find(q => q.Name.Value == configModel1.tableBookMark);
                if (bookmarkStart == null)
                {
                    return;
                }
                var table = bookmarkStart.Parent.Parent.Parent.Parent;
                foreach (var tablemodel in studentInfoModelList)
                {
                    var row = table.Elements<TableRow>().ElementAt(1).Clone() as TableRow;
                    var cells = row.Elements<TableCell>();
                    for (int i = 0; i < cells.Count(); i++)
                    {
                        var cell = cells.ElementAt(i);
                        var tmpPa = cell.Elements<Paragraph>().First();
                        var tmpRuns = tmpPa.Elements<Run>();

                        if (tmpRuns.Count() <= 0)
                        {
                            tmpPa.Remove();
                            tmpPa = new Paragraph(new Run(new Text(" ")));
                            //cell.RemoveAllChildren();
                            cell.Append(tmpPa);

                        }
                        var tmpRun = tmpPa.Elements<Run>().First();
                        var tmpText = tmpRun.Elements<Text>().First();

                        //获取属性值
                        Type type = tablemodel.GetType();
                        string propertyKey = obj[i];
                        System.Reflection.PropertyInfo propertyInfo = type.GetProperty(propertyKey); //获取指定名称的属性
                        object objValue = propertyInfo.GetValue(tablemodel, null);
                        //object objValue = obj[i];
                        if (objValue != null)
                        {
                            tmpText.Text = objValue.ToString().Trim('0',':');
                        }
                        else
                        {
                            tmpText.Text = "-";
                        }
                    }

                    var lastRow = table.Elements<TableRow>().Last();
                    table.InsertAfter(row, lastRow);
                }
                //最后删除startIndex行
                table.Elements<TableRow>().ElementAt(1).Remove();
            }
        }
    }
}

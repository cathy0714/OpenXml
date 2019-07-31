using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
    public class Program
    {
        static string file = @"e:\test\ExcelText.xlsx";
        static string filepath = @"e:\test\Test11.xlsx";
        public static void Main(string[] args)
        {
            ExcelHelper excel = new ExcelHelper(file);
            bool result = excel.CreatExcel(filepath);
            Console.ReadLine();
        }
    }
}

﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml
{
    public class TableDataModel
    {
        public string Property1 { get; set; }
        public string Property2 { get; set; }
        public string Property3 { get; set; }
        public string Property4 { get; set; }
        public string Property5 { get; set; }
    }

    public class ConfigModel
    {
        public int Index { get; set; }
        //开始行索引
        public int StartRowIndex { get; set; }
        //table标签
        public string tableBookMark { get; set; }

        public ConfigModel(string keyid)
        {
            object obj = typeof(BaseConfig).GetField(keyid).GetValue(null);
            if(obj!=null)
            {
                string cigString = obj.ToString();
                string[] array = cigString.Split(',');
                this.Index = Convert.ToInt32(array[0]);
                this.StartRowIndex = Convert.ToInt32(array[1]);
                this.tableBookMark = array[2];
            }
        }
    }

    public class BaseConfig
    {
        public const string table1 = "1,1,表格1";
        public const string table2 = "1,1,表格2";
    }
}

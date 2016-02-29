using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelExtenderDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var src = new DataTable("Data");
            src.Columns.Add("F1");

            src.Columns.Add("C1A", typeof(decimal));
            src.Columns.Add("C1B", typeof(decimal));
            src.Columns.Add("C2A", typeof(decimal));
            src.Columns.Add("C2B", typeof(decimal));
            src.Columns.Add("C3A", typeof(decimal));
            src.Columns.Add("C3B", typeof(decimal));
            src.Columns.Add("C4A", typeof(decimal));
            src.Columns.Add("C4B", typeof(decimal));

            src.Columns.Add("D1", typeof(decimal));
            src.Columns.Add("D2", typeof(decimal));
            src.Columns.Add("D3", typeof(decimal));
            src.Columns.Add("D4", typeof(decimal));
            src.Columns.Add("D5", typeof(decimal));
            src.Columns.Add("D6", typeof(decimal));

            src.Columns.Add("F2", typeof(decimal));

            var rnd = new Random(DateTime.Now.Millisecond);
            DataRow row = null;
            var colIndex = 0;
            for (var i = 0; i < 1000; i++)
            {
                colIndex = 0;
                row = src.NewRow();
                for (var j = src.Columns.Count - 1; j >= 0; j--)
                {
                    row[colIndex++] = rnd.Next(100);
                }

                src.Rows.Add(row);
            }

            Console.WriteLine("Begin processing...");

            WorkbookDesigner designer = new WorkbookDesigner();

            designer.Workbook = new Workbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTemplate1.xls"));//扩展模板ExcelTemplate2.xls

            //ExcelExtender.Extend(designer, new ExcelExtendConfig()//配置自定义扩展
            //{
            //    ColumnRepeatersConfig = new ExcelExtendRepeaterConfig[]
            //    { 
            //        new ExcelExtendRepeaterConfig(){ TagName="a", RepeatCount = 4},//a标签对应的列重复4次
            //        new ExcelExtendRepeaterConfig(){ TagName="b", RepeatCount = 6}//bs标签对应的列重复6次
            //    }
            //});

            designer.SetDataSource(src);//设置列表数据源

            designer.SetDataSource("Other", 1234);//设置变量

            designer.Process();

            var outputFileFullName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("Output/{0}.xls", DateTime.Now.Ticks));
            Directory.CreateDirectory(Path.GetDirectoryName(outputFileFullName));
            designer.Workbook.Save(outputFileFullName);

            Console.WriteLine("输出文件："+outputFileFullName);

            Console.ReadKey();
        }
    }
}

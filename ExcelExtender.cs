using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using Aspose.Cells;
using System.IO;
using System.Web;
using System.Collections;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using ld = ExcelExtenderDemo.Linq.Dynamic;
using System.Linq;
using System.Diagnostics;

namespace ExcelExtenderDemo
{
    /// <summary>
    /// 详细使用文档见： http://km.oa.com/articles/show/266258
    /// </summary>
    public class ExcelExtender
	{
        private const string extendPrefix = "ex::";

        static ExcelExtender()
        {
            ld.ExpressionParser.AddPredefinedType(typeof(ExtenderContext));
            ld.ExpressionParser.AddPredefinedType(typeof(ExtendColumnRepeaterContext));
            ld.ExpressionParser.AddPredefinedType(typeof(ExtendColumnSlot));
            ld.ExpressionParser.AddPredefinedType(typeof(ExtendColumn));
            ld.ExpressionParser.AddPredefinedType(typeof(ExtendCell));
        }

        internal class ExtenderContext
        {
            //public Dictionary<string, ExtendColumnRepeaterContext> Extensions = new Dictionary<string, ExtendColumnRepeaterContext>();

            internal Dictionary<string, TagProcess> TagProcesses = new Dictionary<string, TagProcess>();

            public ExtendColumnRepeaterContext GetRepeater(string tag)
            {
                if (TagProcesses.ContainsKey(tag))
                {
                    var processer = TagProcesses[tag] as ExtendColumnRepeaterTagProcessor;

                    if (processer != null)
                        return processer.Context;
                }

                return null;
            }
        }

        internal class ExtendColumnRepeaterContext
        {
            public ExtendColumnRepeaterContext()
            {
                Slots = new List<ExtendColumnSlot>();
            }

            public string Tag { get; set; }

            public int StartExcelColIndex { get; set; }

            public string StartExcelColName { get { return CellsHelper.ColumnIndexToName(StartExcelColIndex); } }

            public int EndExcelColIndex { get { return StartExcelColIndex + SlotSize * SlotCount - 1; } }

            public string EndExcelColName { get { return CellsHelper.ColumnIndexToName(EndExcelColIndex); } }

            /// <summary>
            /// 总共包含的列数
            /// </summary>
            public int ColumnsCount { get { return SlotSize * SlotCount; } }

            /// <summary>
            /// 一个单元内含多少列
            /// </summary>
            public int SlotSize { get; set; }

            /// <summary>
            /// 重复出现多个个单元
            /// </summary>
            public int SlotCount { get; set; }

            public List<ExtendColumnSlot> Slots { get; set; }//[]

        }

        internal abstract class TagProcess
        {
            private static string EnsurePlaceHolder(ref string placeholder)
            {
                var result = placeholder;
                var braceCount = 0;
                var length = placeholder.Length;
                for (int i = 0; i < length; i++)
                {
                    if (placeholder[i] == '{') braceCount++;
                    else if (placeholder[i] == '}')
                    {
                        braceCount--;

                        if (braceCount == 0)
                        {
                            var start = placeholder.IndexOf(':');
                            placeholder = placeholder.Substring(0, i + 1);
                            result = result.Substring(start + 1, i - start - 1);
                            break;
                        }
                    }
                }

                return result;
            }

            private Regex _regPropRef = new Regex(@"^P\d+\.");
            protected string ReplaceTag(string originalValue, string tagPlaceHolder, List<object[]> argumentsList)
            {
                var tagArgs = EnsurePlaceHolder(ref tagPlaceHolder);//match.Groups["param"].Value

                var args = tagArgs.Split(new char[] { ';' });//, StringSplitOptions.RemoveEmptyEntries 不要remove了，确实有中间的值为空
                var propName = args[0].Trim();
                if (!propName.StartsWith("(") && !_regPropRef.IsMatch(propName)) propName = "P0." + propName;
                var formatString = args.Length > 1 ? args[1].Trim() : "{0}";
                if (string.IsNullOrWhiteSpace(formatString)) formatString = "{0}";
                var spliter = args.Length > 2 ? args[2].Trim() : "";
                var filter = args.Length > 3 ? args[3].Trim() : "";

                string result = "";
                if (argumentsList.Count > 0)
                {
                    var invoker = PraseStringToExress(propName, argumentsList[0]);

                    var results = new List<string>();

                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        switch (filter)
                        {
                            case ":one":
                            case ":first":
                                argumentsList = new List<object[]> { argumentsList.First() };
                                break;
                            case ":last":
                                argumentsList = new List<object[]> { argumentsList.Last() };
                                break;

                            //todo:even
                            //todo:odd
                        }
                    }

                    foreach (var arguments in argumentsList)
                    {
                        results.Add((invoker.DynamicInvoke(arguments) ?? "").ToString());
                    }

                    result = string.Join(spliter, results.Select(r => string.Format(formatString, r)));
                }

                return originalValue.Replace(tagPlaceHolder, result);
            }

            public abstract string Replace(Cell cell, string originalValue, string tagPlaceHolder, ExtenderContext context);
        }

        internal class ExtendColumnRepeaterTagProcessor : TagProcess
        {
            internal ExtendColumnRepeaterContext Context { get; private set; }

            public ExtendColumnRepeaterTagProcessor(ExtendColumnRepeaterContext context)
            {
                Context = context;
                _colSlotMap = new Dictionary<int, ExtendColumnSlot>();
                foreach (var slot in Context.Slots)
                {
                    foreach (var col in slot.Columns)
                    {
                        _colSlotMap.Add(col.ExcelColIndex, slot);
                    }
                }
            }

            public override string Replace(Cell cell, string originalValue, string tagPlaceHolder, ExtenderContext context)
            {
                var regParameter = new Regex(@"\{" + Context.Tag + @":(?<param>.+)\}");//{cp:Columns[1].Name,{0}{{r}}, +}

                var slot = GetSlot(cell);

                var exCell = new ExtendCell(cell);

                if (slot != null)
                {
                    return ReplaceTag(originalValue, tagPlaceHolder, new List<object[]> { new object[] { slot, exCell, Context, context } });
                }
                else
                {
                    return ReplaceTag(originalValue, tagPlaceHolder, Context.Slots.Select(s => new[] { (object)s, exCell, Context, context }).ToList());
                }
            }

            Dictionary<int, ExtendColumnSlot> _colSlotMap = null;
            private ExtendColumnSlot GetSlot(Cell cell)
            {
                if (_colSlotMap.ContainsKey(cell.Column))
                {
                    return _colSlotMap[cell.Column];
                }

                return null;
            }
        }

        internal class ExtendColumnSlot
        {
            public ExtendColumnSlot()
            {
                Columns = new List<ExtendColumn>();
            }

            public int Index { get; set; }

            public int StartExcelColIndex
            {
                get {
                    return Columns[0].ExcelColIndex;//没有列时就抛过异常吧
                }
            }

            public List<ExtendColumn> Columns { get; set; }
        }

        internal class ExtendColumn
        {
            public string Name { get; set; }

            public int ExcelColIndex { get; set; }
        }

        internal class ExtendCell
        {
            private Cell _cell = null;

            //public ExtendCell()
            //{

            //}

            public ExtendCell(Cell cell)
            {
                _cell = cell;
                ColName = cell.Name; 
                ExcelColIndex = cell.Column; 
                ExcelRowIndex = cell.Row;
            }

            public string ColName { get; set; }

            public int ExcelColIndex { get; set; }

            public int ExcelRowIndex { get; set; }

            public string Col(int offset)
            {
                return CellsHelper.ColumnIndexToName(ExcelColIndex + offset);
            }

            public int ColOffset(string start, string end)
            {
                return CellsHelper.ColumnNameToIndex(end) - CellsHelper.ColumnNameToIndex(start);
            }

            public string ColRel(string targetColName, int offset)
            {
                if (string.IsNullOrWhiteSpace(targetColName))
                    return Col(offset);

                return CellsHelper.ColumnIndexToName(CellsHelper.ColumnNameToIndex(targetColName) + offset);
            }

            public string Row(int offset)
            {
                return (ExcelColIndex + offset).ToString();
            }

            public string RowRel(int targetRowIndex, int offset)
            {
                return CellsHelper.ColumnIndexToName(targetRowIndex + offset);
            }

            public string Background(string color)
            {
                _cell.SetStyle(new Style() { BackgroundColor = ColorTranslator.FromHtml(color) });
                return string.Empty;
            }

            public string SetColumnVisible(bool isVisible)
            {
                _cell.Worksheet.Cells.Columns[_cell.Column].IsHidden = isVisible == false;

                return string.Empty;
            }
        }

        internal class ExtendCellTagProcessor : TagProcess
        {
            private Regex _regParameter = new Regex(@"\{" + "cell" + @":(?<param>.+)\}");//{u:col(-2)}
            public override string Replace(Cell cell, string originalValue, string tagPlaceHolder, ExtenderContext context)
            {
                return base.ReplaceTag(originalValue, tagPlaceHolder, new List<object[]> { new object[] { new ExtendCell(cell), context } });
            }
        }

        /// <summary>
        /// <para>空白处标记 ex::cp(cr:2)  表示tag为cp的扩展，cr:2以当前列和后面1(2-1)列为一个整体，重复复制扩展；cell 是被公用方法占用</para>
        /// <para>{cp:Index, {0}{{r}}, spliter, filter} {访问的属性, 显示的格式[可选], 多值合并时的分隔符[可选], 过滤器[可选]}</para>
        /// <para>{cp:Index}：出现在重复列时，取当前ExtendColumnSlot中的值</para>
        /// <para>{cp:Index}：出现在非重复列时，取所有所有ExtendColumnSlot中的值后合并为一个值</para>
        /// <para>{cp:P1.ColumnsCount}：默认访问的是ExtendColumnSlot对象，可以使用P1.前缀访问ExtendCell对象，可以使用P2.前缀访问ExtendColumnRepeaterContext对象，可以使用P3.前缀访问ExtendContext对象</para>
        /// <para>{cell:Col(-1)}：访问ExtendCell中的Col方法，返回左侧一列的列名</para>
        /// <para>如：</para>
        /// <para>=$X{cp:Index}Name -> 替换为当前组的索引号(从1开始)，即：=$X1Name 或 =$X2Name ...</para>
        /// <para>=$L{r}+$N{r}+{cp:Columns[1].Name,${0}{{r}}, +} -> 替换为当前组中第二列的列名并按格式做合并，即：=$L{r}+$N{r}+$A{r}+$C{r}+$E{r} ... </para>
        /// <para>{a:Columns[0].Name,,,:last} -> 取所有循环组中最后一组中的第一列的列名</para>
        /// </summary>
        public static void Extend(WorkbookDesigner wbd, ExcelExtendConfig config = null)
        {
            var sheet = wbd.Workbook.Worksheets[0];
            var context = new ExtenderContext();

            //extension
            var cfgCell = sheet.Cells.FindStringStartsWith(extendPrefix, null);//todo: replce with Find
            while (cfgCell != null)
            {
                var extTagCfg = cfgCell.Value.ToString().Replace(extendPrefix, "");//ex::cp(cr:2[,...])
                cfgCell.PutValue(null);//清空界面上的配置内容

                var regTemp = new Regex(@"(?<tag>\w*?)(\((?<cfg>.+)\))");
                var match = regTemp.Match(extTagCfg);

                if (match.Success)
                {
                    var tag = match.Groups["tag"].Value;
                    var dicRepeaterConfig = config == null || config.ColumnRepeatersConfig == null ? new Dictionary<string, ExcelExtendRepeaterConfig>() : config.ColumnRepeatersConfig.ToDictionary(c => c.TagName);
                    //regTemp = new Regex(@"(?<op>\w+)(:(?<p>.*),?)?");
                    //match = regTemp.Match(match.Groups["cfg"].Value);//cr:2[,...]
                    foreach (var extension in match.Groups["cfg"].Value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var tuple = extension.Split(':');
                        var exCfg = tuple.Length > 1 ? tuple[1].Trim() : null;
                        switch (tuple[0].Trim().ToLower())
                        {
                            case "cr"://列重复
                                var repeatCount = 0;//
                                if (dicRepeaterConfig.ContainsKey(tag))
                                    repeatCount = dicRepeaterConfig[tag].RepeatCount;
                                else if( dicRepeaterConfig.Count ==1)
                                    repeatCount = dicRepeaterConfig.First().Value.RepeatCount;

                                var cr = ExtendExcelColumnRepeater(sheet, tag, cfgCell.Column, exCfg, repeatCount);
                                //context.Extensions.Add(tag, cr.Context);
                                context.TagProcesses.Add(tag, cr);
                                break;
                            case "rr"://行重复
                                //todo:行重复
                                break;
                        }
                    }
                }

                cfgCell = sheet.Cells.FindStringStartsWith(extendPrefix, cfgCell);
            }

            //replace all tag
            context.TagProcesses.Add("cell", new ExtendCellTagProcessor());

            ReplaceAllTags(sheet, context);
        }

        private static ExtendColumnRepeaterTagProcessor ExtendExcelColumnRepeater(Worksheet sheet, string tag, int startColIndex, string conf, int crCount)
        {
            if (crCount < 0) crCount = 1;

            int colCount = 1;
            if (!string.IsNullOrWhiteSpace(conf))
            {
                colCount = int.Parse(conf);
            }

            var context = new ExtendColumnRepeaterContext()
            {
                Tag = tag,
                StartExcelColIndex = startColIndex,
                SlotSize = colCount,
                SlotCount = crCount
            };

            return ExtendExcelColumnRepeater(sheet, context);
        }

        private static ExtendColumnRepeaterTagProcessor ExtendExcelColumnRepeater(Worksheet sheet, ExtendColumnRepeaterContext context)
        {
            var currentColIndex = context.StartExcelColIndex;

            FillSlotColumns(context, 1, currentColIndex);//将已存在的配置列加入到Context中

            if (context.SlotCount > 0)
            {
                for (var index = 2; index <= context.SlotCount; index++)
                {
                    sheet.Cells.InsertColumns(currentColIndex + context.SlotSize, context.SlotSize);
                    sheet.Cells.CopyColumns(sheet.Cells, currentColIndex, currentColIndex + context.SlotSize, context.SlotSize);
                    FillSlotColumns(context, index, currentColIndex + context.SlotSize);
                    currentColIndex += context.SlotSize;
                }

                SwiftColumnInFormula(sheet, context);
            }
            else//重复列数小于等于零，隐藏原配置列
            {
                for (var offset = 0; offset < context.SlotSize; offset++)
                {
                    sheet.Cells.Columns[currentColIndex + offset].IsHidden = true;
                }
            }

            return new ExtendColumnRepeaterTagProcessor(context);
        }

        //修改公式：主要是列名移动
        private static void SwiftColumnInFormula(Worksheet sheet, ExtendColumnRepeaterContext context)
        {
            var regParameter = new Regex(@"\$(?<col>[A-Z]{1,2})");//{u:col(-2)}{}
            var newColsCount = context.ColumnsCount - context.SlotSize;//总数-原本存在的列（被复制）
            var oldRange = new Tuple<int, int>(context.StartExcelColIndex, context.StartExcelColIndex + context.SlotSize - 1);
            foreach (Cell cell in sheet.Cells)
            {
                var value = cell.Formula ?? cell.StringValue;
                if (value.StartsWith("&=$")) continue;//aspose 的变量内容与列表无法区别，所以先忽略

                Match match = null;
                while ((match = regParameter.Match(value, match == null ? 0 : match.Index + 1)).Success)
                {
                    var gCol = match.Groups["col"];
                    var colName = gCol.Value;
                    var refCol = CellsHelper.ColumnNameToIndex(colName);
                    if (refCol > oldRange.Item2)
                    {
                        colName = CellsHelper.ColumnIndexToName(refCol + newColsCount);
                        value = value.Remove(gCol.Index, gCol.Value.Length);
                        value = value.Insert(gCol.Index, colName);
                        cell.Value = value;
                    }
                }
            }
        }

        private static void FillSlotColumns(ExtendColumnRepeaterContext context, int index, int startIndex)
        {
            var slot = new ExtendColumnSlot() { Index = index };
            for (var offset = 0; offset < context.SlotSize; offset++)
            {
                //var col = sheet.Cells.Columns[startColIndex + offset];
                slot.Columns.Add(new ExtendColumn() { ExcelColIndex = startIndex + offset, Name = CellsHelper.ColumnIndexToName(startIndex + offset) });
            }

            context.Slots.Add(slot);
        }

        private static void ReplaceAllTags(Worksheet sheet, ExtenderContext context)
        {
            var regParameter = new Regex(@"\{(?<tag>\w+):(?<param>.+)\}");//{u:col(-2)}{}
            foreach (Cell cell in sheet.Cells)//todo: replce with Find
            {
                Match match = null;
                var value = cell.Formula ?? cell.StringValue;
                while ((match = regParameter.Match(value)).Success)
                {
                    var tag = match.Groups["tag"].Value;
                    if (context.TagProcesses.ContainsKey(tag) && context.TagProcesses[tag] != null)
                    {
                        cell.Value = context.TagProcesses[tag].Replace(cell, value, match.Value, context);
                    }

                    value = cell.Formula ?? cell.StringValue;
                }
            }
        }
        
        private static Delegate PraseStringToExress(string express, params object[] args)
        {//todo:cache expression result
            var ps = new ParameterExpression[args.Length];
            for (var i = 0; i < args.Length; i++)
            {
                ps[i] = Expression.Parameter(args[i].GetType(), "P" + i);
            }

            var e = ld.DynamicExpression.ParseLambda(ps, null, express, args);

            return e.Compile();
        }
    }

    public class ExcelExtendConfig
    {
        public ExcelExtendRepeaterConfig[] ColumnRepeatersConfig { get; set; }
    }

    public class ExcelExtendRepeaterConfig
    {
        public string TagName { get; set; }

        public int RepeatCount { get; set; }
    }
}

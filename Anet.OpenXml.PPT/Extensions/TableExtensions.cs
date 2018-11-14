using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Anet.OpenXml.PPT
{
    public static class TableExtensions
    {
        /// <summary>
        /// 设置表格列宽
        /// </summary>
        public static GridColumn SetColumnWidth(this Table table, int col, long width)
        {
            var gridColumn = table.TableGrid.Elements<GridColumn>().ElementAt(col);
            gridColumn.Width = width;
            return gridColumn;
        }

        /// <summary>
        /// 设置表格行高
        /// </summary>
        public static TableRow SetRowHeight(this Table table, int row, long height)
        {
            var tableRow = table.Elements<TableRow>().ElementAt(row);
            tableRow.Height = height;
            return tableRow;
        }

        /// <summary>
        /// 替换表格数据
        /// </summary>
        /// <param name="srcData">数据源，只有应用了Display特性的属性才会被写入</param>
        /// <param name="startRow">起始行（从0开始）</param>
        /// <param name="startCol">取始列（从0开始）</param>
        /// <param name="callback">
        /// 自定义处理回调。
        /// 输入参数：行号，列号，当前单元格；
        /// 输出参数：是否已自定义处理，若已自定义处理则不再会自动填值
        /// </param>
        public static void ReplaceData<T>(this Table table, List<T> srcData, int startRow, int startCol,
            Func<int, int, TableCell, T, bool> callback) where T : new()
        {
            // 获取具排序好的有Display特性的属性
            var props = GetSortedProps<T>();
            if (props.Count() == 0)
                return;

            var tableRows = table.Elements<TableRow>();

            for (int i = 0; i < tableRows.Count(); i++)
            {
                if (i < startRow)
                    continue;

                // 数据源行数不足则返回
                if (i - startRow > srcData.Count - 1)
                    return;

                var dataItem = srcData[i - startRow];
                var tableCells = tableRows.ElementAt(i).Elements<TableCell>();
                for (int j = 0; j < tableCells.Count(); j++)
                {
                    if (j < startCol)
                        continue;

                    var tableCell = tableCells.ElementAt(j);
                    var theProp = props.ElementAt(j - startCol);

                    // 如果自定义回调处理该单元格返回返回了True，则不自动填写内容
                    if (callback != null && callback(i, j, tableCell, dataItem))
                        continue;

                    var propVal = (theProp.GetValue(dataItem) ?? "").ToString();

                    var text = tableCell.Descendants<Text>().FirstOrDefault();
                    if (text == null)
                    {
                        // 如果所给值不为为则创建一个Text填内容
                        if (!string.IsNullOrEmpty(propVal))
                        {
                            text = new Text { Text = propVal };
                            tableCell.Append(text);
                        }
                    }
                    else
                    {
                        text.Text = propVal;
                    }
                }
            }
        }

        /// <summary>
        /// 获取按顺序排列的属性
        /// </summary>
        private static IEnumerable<PropertyInfo> GetSortedProps<T>()
        {
            var dic = new SortedDictionary<int, PropertyInfo>();
            var props = typeof(T).GetProperties();
            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<DisplayAttribute>();
                if (attr == null)
                    continue;
                dic.Add(attr.Order, prop);
            }
            return dic.Values;
        }
    }
}

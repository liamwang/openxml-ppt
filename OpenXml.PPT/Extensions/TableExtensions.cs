using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OpenXml.PPT
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
        /// 获取列表中间的元素
        /// </summary>
        public static T GetMiddleItem<T>(this IEnumerable<T> list)
        {
            var mid = (int)Math.Ceiling(list.Count() / 2f) - 1;
            return list.ElementAt(mid);
        }

        /// <summary>
        /// 填充表格数据（根据需要动态增加行和列）
        /// 模板中必须有一个至少1行1列的表格
        /// </summary>
        public static void FillData(this Table table, List<string[]> data)
        {
            if (data == null || data.Count == 0)
                return;

            // 目标列数
            var columnCount = data[0].Length;

            // 模板列信息
            var tableGrid = table.Elements<TableGrid>().First();
            var gridColumns = tableGrid.Elements<GridColumn>();

            // 模板行信息
            var tableRows = table.Elements<TableRow>();

            // 根据需要增加列
            if (gridColumns.Count() < columnCount)
            {
                var refColumn = gridColumns.GetMiddleItem();
                for (var i = gridColumns.Count(); i < columnCount; i++)
                {
                    var newColumn = refColumn.CloneNode(true);
                    tableGrid.InsertAfter(newColumn, refColumn);
                    foreach (var tableRow in tableRows)
                    {
                        var refCell = tableRow.Elements<TableCell>().GetMiddleItem();
                        var newCell = refCell.CloneNode(true);
                        tableRow.InsertAfter(newCell, refCell);
                    }
                }
            }

            // 填充行
            for (int row = 0; row < data.Count; row++)
            {
                TableRow tableRow = null;

                // 如果表格行数不足，则复制上一行插入
                if (row > tableRows.Count() - 1)
                {
                    tableRow = tableRows.ElementAt(row - 1).CloneNode(true) as TableRow;
                    table.Append(tableRow);
                }
                else
                {
                    tableRow = tableRows.ElementAt(row);
                }

                // 填充单元格
                var tableCells = tableRow.Elements<TableCell>();
                for (int col = 0; col < tableCells.Count(); col++)
                {
                    var value = data[row][col];
                    var tableCell = tableCells.ElementAt(col);
                    var text = tableCell.Descendants<Text>().FirstOrDefault();
                    if (text == null)
                    {
                        // 如果所给值不为空则创建一个Text填内容
                        if (value != null)
                        {
                            text = new Text { Text = value };
                            tableCell.Append(text);
                        }
                    }
                    else
                    {
                        text.Text = value;
                    }
                }
            }
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
            if (srcData == null || srcData.Count == 0)
                return;

            // 获取具排序好的有Display特性的属性
            var props = TypeUtil.GetOrderedDisplayProps<T>();
            if (props.Count() == 0)
                return;

            var tableRows = table.Elements<TableRow>().Skip(startRow);
            if (tableRows.Count() == 0)
                return;

            for (int row = 0; row < srcData.Count; row++)
            {
                TableRow tableRow = null;

                // 如果表格行数不足，则复制上一行插入
                if (row > tableRows.Count() - 1)
                {
                    tableRow = tableRows.ElementAt(row - 1).CloneNode(true) as TableRow;
                    table.Append(tableRow);
                }
                else
                {
                    tableRow = tableRows.ElementAt(row);
                }

                ReplaceRowData(tableRow, srcData[row], props, row, startCol, callback);
            }
        }

        /// <summary>
        /// 替换表格某行数据
        /// </summary>
        /// <param name="srcData">数据源，只有应用了Display特性的属性才会被写入</param>
        /// <param name="startCol">取始列（从0开始）</param>
        /// <param name="callback">
        /// 自定义处理回调。
        /// 输入参数：行号，列号，当前单元格；
        /// 输出参数：是否已自定义处理，若已自定义处理则不再会自动填值
        /// </param>
        private static TableRow ReplaceRowData<T>(TableRow row, T model, IEnumerable<PropertyInfo> props, int rowIndex, int startCol,
            Func<int, int, TableCell, T, bool> callback) where T : new()
        {
            var tableCells = row.Elements<TableCell>();
            for (int col = 0; col < tableCells.Count(); col++)
            {
                if (col < startCol)
                    continue;

                var tableCell = tableCells.ElementAt(col);
                var theProp = props.ElementAt(col - startCol);

                // 如果自定义回调处理该单元格返回返回了True，则不自动填写内容
                if (callback != null && callback(rowIndex, col, tableCell, model))
                    continue;

                var propVal = (theProp.GetValue(model) ?? "").ToString();

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
            return row;
        }
    }
}

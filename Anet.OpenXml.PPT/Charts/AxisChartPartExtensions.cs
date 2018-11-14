using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;

namespace Anet.OpenXml.PPT.Charts
{
    public static class AxisChartPartExtensions
    {
        ///// <summary>
        ///// 修改条形图数据
        ///// </summary>
        ///// <param name="model">条形图数据</param>
        //public static void ResetBarChartSeriesData(this ChartPart chartPart, List<ChartSeriesModel> data, IEnumerable<string> categories)
        //{
        //    // 图表不能超过6个系列（SchemeColorValues.Accent1~6）
        //    if (data.Count > 16)
        //    {
        //        throw new Exception("图表不能超过6个系列！");
        //    }

        //    var barChart = chartPart.ChartSpace
        //        .GetFirstChild<C.Chart>()
        //        .GetFirstChild<PlotArea>()
        //        .GetFirstChild<BarChart>();

        //    // 移除所有系列
        //    barChart.RemoveAllChildren<BarChartSeries>();

        //    for (int i = 0; i < data.Count; i++)
        //    {
        //        var item = data[i];

        //        var barChartSeries = new BarChartSeries();
        //        barChartSeries.Append(new Index() { Val = (uint)i });
        //        barChartSeries.Append(new Order() { Val = (uint)i });

        //        if (!string.IsNullOrEmpty(item.Title))
        //        {
        //            barChartSeries.Append(new SeriesText(
        //                new StringReference(
        //                    // new Formula { Text = "" }
        //                    new StringCache(
        //                        new PointCount { Val = 1U },
        //                        new StringPoint(new NumericValue { Text = item.Title }) { Index = 0U }
        //                    )
        //                )
        //            ));
        //        }

        //        barChartSeries.Append(new ChartShapeProperties(
        //            new SolidFill(new SchemeColor { Val = GetAccentColor(i) }),
        //            new Outline(new NoFill()),
        //            new EffectList()
        //        ));

        //        barChartSeries.Append(new InvertIfNegative { Val = false });

        //        // 分类
        //        if (categories != null && categories.Count() > 0)
        //        {
        //            barChartSeries.Append(new CategoryAxisData().ChangeData(categories));
        //        }

        //        // 数据
        //        barChartSeries.Append(GenerateValues(item.Data, item.FormatCode));

        //        barChart.Append(barChartSeries);
        //    }
        //}

        ///// <summary>
        ///// 修改拆线图数据
        ///// </summary>
        ///// <param name="model">条形图数据</param>
        //public static void ResetLineChartSeriesData(this ChartPart chartPart, List<ChartSeriesModel> data, IEnumerable<string> categories)
        //{
        //    // 图表不能超过6个系列（SchemeColorValues.Accent1~6）
        //    if (data.Count > 16)
        //    {
        //        throw new Exception("图表不能超过6个系列！");
        //    }

        //    var lineChart = chartPart.ChartSpace
        //        .GetFirstChild<C.Chart>()
        //        .GetFirstChild<PlotArea>()
        //        .GetFirstChild<LineChart>();

        //    // 移除所有系列
        //    lineChart.RemoveAllChildren<LineChartSeries>();

        //    for (int i = 0; i < data.Count; i++)
        //    {
        //        var item = data[i];

        //        var chartSeries = new LineChartSeries();
        //        chartSeries.Append(new Index() { Val = (uint)i + 2 });
        //        chartSeries.Append(new Order() { Val = (uint)i + 2 });

        //        if (!string.IsNullOrEmpty(item.Title))
        //        {
        //            chartSeries.Append(new SeriesText(
        //                new StringReference(
        //                    // new Formula { Text = "" }
        //                    new StringCache(
        //                        new PointCount { Val = 1U },
        //                        new StringPoint(new NumericValue { Text = item.Title }) { Index = 0U }
        //                    )
        //                )
        //            ));
        //        }

        //        var chartShapeProperties = new ChartShapeProperties();
        //        var outline = new Outline() { Width = 28575, CapType = LineCapValues.Round };
        //        outline.Append(new SolidFill(new SchemeColor() { Val = GetAccentColor(i) }));
        //        outline.Append(new Round());
        //        chartShapeProperties.Append(outline);
        //        chartShapeProperties.Append(new EffectList());
        //        chartSeries.Append(chartShapeProperties);

        //        var marker = new Marker();
        //        var chartShapeProperties2 = new ChartShapeProperties();
        //        var solidFill2 = new SolidFill(new SchemeColor() { Val = GetAccentColor(i) });
        //        var outline2 = new Outline() { Width = 9525 };
        //        outline2.Append(new SolidFill(new SchemeColor() { Val = GetAccentColor(i) }));
        //        chartShapeProperties2.Append(solidFill2);
        //        chartShapeProperties2.Append(outline2);
        //        chartShapeProperties2.Append(new EffectList());

        //        marker.Append(new Symbol() { Val = C.MarkerStyleValues.Circle });
        //        marker.Append(new Size() { Val = 5 });
        //        marker.Append(chartShapeProperties2);
        //        chartSeries.Append(marker);

        //        // 分类
        //        if (categories != null && categories.Count() > 0)
        //        {
        //            chartSeries.Append(new CategoryAxisData().ChangeData(categories));
        //        }

        //        // 数据
        //        chartSeries.Append(GenerateValues(item.Data, item.FormatCode));

        //        lineChart.Append(chartSeries);
        //    }
        //}

        /// <summary>
        /// 修改条形图数据
        /// </summary>
        /// <param name="model">条形图数据</param>
        public static void ResetChartSeriesData(this ChartPart chartPart, List<ChartSeriesModel> data, IEnumerable<string> categories)
        {
            // 图表不能超过6个系列（SchemeColorValues.Accent1~6）
            if (data.Count > 6)
            {
                throw new Exception("图表不能超过6个系列！");
            }

            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>().GetFirstChild<PlotArea>();
            var barChart = plotArea.GetFirstChild<BarChart>();
            var lineChart = plotArea.GetFirstChild<LineChart>();

            // 移除所有系列
            if (barChart != null)
                barChart.RemoveAllChildren<BarChartSeries>();
            if (lineChart != null)
                lineChart.RemoveAllChildren<LineChartSeries>();

            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];
                if (item.ChartType == ChartType.BarChart)
                {
                    var chartSeries = GenerateBarChartSeries(item, categories, i);
                    barChart.Append(chartSeries);
                }
                else if (item.ChartType == ChartType.LineChart)
                {
                    var chartSeries = GenerateLineChartSeries(item, categories, i);
                    lineChart.Append(chartSeries);
                }
            }
        }

        private static LineChartSeries GenerateLineChartSeries(ChartSeriesModel model, IEnumerable<string> categories, int index)
        {
            var chartSeries = new LineChartSeries();
            chartSeries.Append(new Index() { Val = (uint)index + 2 });
            chartSeries.Append(new Order() { Val = (uint)index + 2 });

            if (!string.IsNullOrEmpty(model.Title))
            {
                chartSeries.Append(new SeriesText(
                    new StringReference(
                        // new Formula { Text = "" }
                        new StringCache(
                            new PointCount { Val = 1U },
                            new StringPoint(new NumericValue { Text = model.Title }) { Index = 0U }
                        )
                    )
                ));
            }

            var chartShapeProperties = new ChartShapeProperties();
            var outline = new Outline() { Width = 28575, CapType = LineCapValues.Round };
            outline.Append(new SolidFill(new SchemeColor() { Val = GetAccentColor(index) }));
            outline.Append(new Round());
            chartShapeProperties.Append(outline);
            chartShapeProperties.Append(new EffectList());
            chartSeries.Append(chartShapeProperties);

            var marker = new Marker();
            var chartShapeProperties2 = new ChartShapeProperties();
            var solidFill2 = new SolidFill(new SchemeColor() { Val = GetAccentColor(index) });
            var outline2 = new Outline() { Width = 9525 };
            outline2.Append(new SolidFill(new SchemeColor() { Val = GetAccentColor(index) }));
            chartShapeProperties2.Append(solidFill2);
            chartShapeProperties2.Append(outline2);
            chartShapeProperties2.Append(new EffectList());

            marker.Append(new Symbol() { Val = MarkerStyleValues.Circle });
            marker.Append(new Size() { Val = 5 });
            marker.Append(chartShapeProperties2);
            chartSeries.Append(marker);

            // 分类
            if (categories != null && categories.Count() > 0)
            {
                chartSeries.Append(new CategoryAxisData().ChangeData(categories));
            }

            // 数据
            chartSeries.Append(GenerateValues(model.Data, model.FormatCode));

            return chartSeries;
        }

        private static BarChartSeries GenerateBarChartSeries(ChartSeriesModel item, IEnumerable<string> categories, int index)
        {
            var barChartSeries = new BarChartSeries();
            barChartSeries.Append(new Index() { Val = (uint)index });
            barChartSeries.Append(new Order() { Val = (uint)index });

            if (!string.IsNullOrEmpty(item.Title))
            {
                barChartSeries.Append(new SeriesText(
                    new StringReference(
                        // new Formula { Text = "" }
                        new StringCache(
                            new PointCount { Val = 1U },
                            new StringPoint(new NumericValue { Text = item.Title }) { Index = 0U }
                        )
                    )
                ));
            }

            barChartSeries.Append(new ChartShapeProperties(
                new SolidFill(new SchemeColor { Val = GetAccentColor(index) }),
                new Outline(new NoFill()),
                new EffectList()
            ));

            barChartSeries.Append(new InvertIfNegative { Val = false });

            // 分类
            if (categories != null && categories.Count() > 0)
            {
                barChartSeries.Append(new CategoryAxisData().ChangeData(categories));
            }

            // 数据
            barChartSeries.Append(GenerateValues(item.Data, item.FormatCode));

            return barChartSeries;
        }

        private static Values GenerateValues(double[] data, string formatCode = null)
        {
            var values = new Values();
            var numberReference = new NumberReference();
            //numberReference.Append(new Formula { Text = "" });

            var numberingCache = new NumberingCache();
            if (!string.IsNullOrEmpty(formatCode))
            {
                numberingCache.Append(new FormatCode { Text = formatCode });
            }
            numberingCache.Append(new PointCount { Val = (uint)data.Length });
            for (int i = 0; i < data.Length; i++)
            {
                var numericPoint = new NumericPoint { Index = (uint)i };
                numericPoint.Append(new NumericValue { Text = data[i].ToString() });
                numberingCache.Append(numericPoint);
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);

            return values;
        }

        private static SchemeColorValues GetAccentColor(int index)
        {
            return (SchemeColorValues)Enum.Parse(typeof(SchemeColorValues), "Accent" + (index + 1));
        }
    }
}

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
        /// <summary>
        /// 修改条形图数据
        /// </summary>
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

            // 创建模式

            // 移除所有系列
            if (barChart != null)
                barChart.RemoveAllChildren<BarChartSeries>();
            if (lineChart != null)
                lineChart.RemoveAllChildren<LineChartSeries>();

            for (int i = 0; i < data.Count; i++)
            {
                var item = data[i];
                if (barChart != null && item.ChartType == ChartType.BarChart)
                {
                    var chartSeries = GenerateBarChartSeries(item, categories, i);
                    barChart.Append(chartSeries);
                }
                else if (lineChart != null && item.ChartType == ChartType.LineChart)
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
            outline.Append(new SolidFill(new SchemeColor() { Val = ChartUtil.GetAccentColor(index) }));
            outline.Append(new Round());
            chartShapeProperties.Append(outline);
            chartShapeProperties.Append(new EffectList());
            chartSeries.Append(chartShapeProperties);

            var marker = new Marker();
            var chartShapeProperties2 = new ChartShapeProperties();
            var solidFill2 = new SolidFill(new SchemeColor() { Val = ChartUtil.GetAccentColor(index) });
            var outline2 = new Outline() { Width = 9525 };
            outline2.Append(new SolidFill(new SchemeColor() { Val = ChartUtil.GetAccentColor(index) }));
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
            chartSeries.Append(ChartUtil.GenerateValues(model.Data, model.FormatCode));

            var dataLables = GenerateDataLabels(chartSeries);
            chartSeries.Append(dataLables);

            return chartSeries;
        }

        private static BarChartSeries GenerateBarChartSeries(ChartSeriesModel item, IEnumerable<string> categories, int index)
        {
            var chartSeries = new BarChartSeries();
            chartSeries.Append(new Index() { Val = (uint)index });
            chartSeries.Append(new Order() { Val = (uint)index });

            if (!string.IsNullOrEmpty(item.Title))
            {
                chartSeries.Append(new SeriesText(
                    new StringReference(
                        // new Formula { Text = "" }
                        new StringCache(
                            new PointCount { Val = 1U },
                            new StringPoint(new NumericValue { Text = item.Title }) { Index = 0U }
                        )
                    )
                ));
            }

            chartSeries.Append(new ChartShapeProperties(
                new SolidFill(new SchemeColor { Val = ChartUtil.GetAccentColor(index) }),
                new Outline(new NoFill()),
                new EffectList()
            ));

            chartSeries.Append(new InvertIfNegative { Val = false });

            // 分类
            if (categories != null && categories.Count() > 0)
            {
                chartSeries.Append(new CategoryAxisData().ChangeData(categories));
            }

            // 数据
            chartSeries.Append(ChartUtil.GenerateValues(item.Data, item.FormatCode));

            var dataLables = GenerateDataLabels(chartSeries);
            chartSeries.Append(dataLables);

            return chartSeries;
        }

        private static DataLabels GenerateDataLabels(OpenXmlElement chartSeries)
        {
            var dataLabels = new C.DataLabels();

            dataLabels.Append(new C.NumberingFormat() { FormatCode = "#,##0_);[Red]\\(#,##0\\)", SourceLinked = false });

            dataLabels.Append(new C.ShowLegendKey() { Val = false });
            dataLabels.Append(new C.ShowValue() { Val = true });
            dataLabels.Append(new C.ShowCategoryName() { Val = false });
            dataLabels.Append(new C.ShowSeriesName() { Val = false });
            dataLabels.Append(new C.ShowPercent() { Val = false });
            dataLabels.Append(new C.ShowBubbleSize() { Val = false });
            dataLabels.Append(new C.ShowLeaderLines() { Val = false });

            return dataLabels;
        }
    }
}

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXml.PPT.Charts
{
    public static class ChartPartExtensions
    {
        /// <summary>
        /// 修改标题
        /// </summary>
        /// <param name="title">标题</param>
        public static void ChangeTitle(this ChartPart chartPart, string title)
        {
            if (title == null) return;

            var titleEl = chartPart.ChartSpace.Descendants<Title>().FirstOrDefault();

            if (titleEl == null) return;

            var text = titleEl.Descendants<Text>().FirstOrDefault();
            if (text != null)
            {
                text.Text = title;
            }
        }

        /// <summary>
        /// 修改坐标辆分类
        /// </summary>
        /// <param name="categories">分类数据</param>
        public static void ChangeCategoryAxisData(this ChartPart chartPart, IEnumerable<string> categories)
        {
            if (categories == null) return;

            // 注：每个系列都有各自的一样的坐标轴分类

            var categoryAxisDataElements = chartPart.ChartSpace.Descendants<CategoryAxisData>();
            if (categoryAxisDataElements.Count() == 0)
                return;

            foreach (var categoryAxisData in categoryAxisDataElements)
            {
                categoryAxisData.ChangeData(categories);
            }
        }

        /// <summary>
        /// 修改条形图数据
        /// </summary>
        /// <param name="model">条形图数据</param>
        public static void ChangeAxisChartData(this ChartPart chartPart, AxisChartModel model)
        {
            // 标题
            chartPart.ChangeTitle(model.Title);

            //// 坐标轴分类
            //chartPart.ChangeCategoryAxisData(model.Categories);

            // 坐标轴最大值
            if (model.MaxAxisValue.HasValue)
            {
                var maxAxisValue = chartPart.ChartSpace.Descendants<MaxAxisValue>().FirstOrDefault();
                if (maxAxisValue != null)
                {
                    maxAxisValue.Val = model.MaxAxisValue.Value;
                }
            }

            // 坐标轴步长单位
            if (model.MajorUnit.HasValue)
            {
                var majorUnit = chartPart.ChartSpace.Descendants<MajorUnit>().FirstOrDefault();
                if (majorUnit != null)
                {
                    majorUnit.Val = model.MajorUnit.Value;
                }
            }

            chartPart.ResetChartSeriesData(model.SeriesList, model.Categories);

            //// 表格数据
            //if (model.SeriesList == null) return;

            //var chart = chartPart.ChartSpace.Descendants<BarChart>().First();
            //var serieses = chart.Elements<BarChartSeries>();

            //if (model.SeriesList.Count != serieses.Count())
            //    throw new Exception("所给数据的系列数和模板条型图不一致。");

            //for (int i = 0; i < model.SeriesList.Count; i++)
            //{
            //    var series = serieses.ElementAt(i);
            //    var seriesData = model.SeriesList[i];

            //    // 替换系列标题
            //    var seriesText = series.Elements<SeriesText>().FirstOrDefault();
            //    if (seriesText != null && !string.IsNullOrWhiteSpace(seriesData.Title))
            //    {
            //        var numericValue = seriesText.Descendants<NumericValue>().First();
            //        numericValue.Text = seriesData.Title;
            //    }

            //    // 替换系列数据
            //    var numberingCache = series.Descendants<NumberingCache>().First();
            //    var numericPoints = numberingCache.Elements<NumericPoint>();

            //    if (seriesData.Data.Length != numericPoints.Count())
            //        throw new Exception("条形图数据长度和模板不一致。");

            //    for (int j = 0; j < seriesData.Data.Length; j++)
            //    {
            //        var numericValue = numericPoints.ElementAt(j).GetFirstChild<NumericValue>();
            //        numericValue.Text = seriesData.Data[j].ToString();
            //    }
            //}
        }
    }
}

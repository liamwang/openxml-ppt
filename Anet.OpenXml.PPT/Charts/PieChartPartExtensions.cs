using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Anet.OpenXml.PPT.Charts
{
    public static class PieChartPartExtensions
    {
        /// <summary>
        /// 修改饼图数据
        /// </summary>
        /// <param name="model">饼图数据</param>
        public static void ChangePieData(this ChartPart chartPart, PieChartModel model)
        {
            // 标题
            chartPart.ChangeTitle(model.Title);

            // 分类
            chartPart.ChangeCategoryAxisData(model.Categories);

            var chart = chartPart.ChartSpace.Descendants<PieChart>().FirstOrDefault();

            var numberingCache = chart.Descendants<NumberingCache>().First();
            var numericPoints = numberingCache.Elements<NumericPoint>();

            if (model.Data.Count() != numericPoints.Count())
                throw new Exception("饼图数据长度和模板不匹配。");

            for (int i = 0; i < numericPoints.Count(); i++)
            {
                var numericValue = numericPoints.ElementAt(i).GetFirstChild<NumericValue>();
                numericValue.Text = model.Data.ElementAt(i).ToString();
            }
        }
    }
}

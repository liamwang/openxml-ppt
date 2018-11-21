using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Anet.OpenXml.PPT.Charts
{
    public static class PieChartPartExtensions
    {
        /// <summary>
        /// 修改饼图数据
        /// </summary>
        /// <param name="model">条形图数据</param>
        public static void ResetPieChartData(this ChartPart chartPart, PieChartModel model)
        {
            // 图表不能超过6个系列（SchemeColorValues.Accent1~6）
            if (model.Data.Count() > 6)
            {
                throw new Exception("图表不能超过6个系列！");
            }

            // 标题
            if (!string.IsNullOrEmpty(model.Title))
            {
                chartPart.ChangeTitle(model.Title);
            }

            var pieChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()
                .GetFirstChild<PlotArea>().GetFirstChild<PieChart>();

            var pieChartSeries = pieChart.GetFirstChild<PieChartSeries>();

            // 区块
            pieChartSeries.RemoveAllChildren<DataPoint>();
            for (int i = 0; i < model.Data.Count(); i++)
            {
                var dataPoint = new DataPoint();
                var chartShapeProperties = new ChartShapeProperties();

                var solidFill1 = new SolidFill();
                var schemeColor1 = new SchemeColor() { Val = ChartUtil.GetAccentColor(i) };
                solidFill1.Append(schemeColor1);

                var outline2 = new Outline() { Width = 19050 };
                var solidFill2 = new SolidFill();
                var schemeColor2 = new SchemeColor() { Val = A.SchemeColorValues.Light1 };
                solidFill2.Append(schemeColor2);
                outline2.Append(solidFill2);

                chartShapeProperties.Append(solidFill1);
                chartShapeProperties.Append(outline2);
                chartShapeProperties.Append(new EffectList());

                dataPoint.Append(new Index() { Val = (uint)i });
                dataPoint.Append(new Bubble3D() { Val = false });
                dataPoint.Append(chartShapeProperties);
                //dataPoint.Append(extensionList1);

                pieChartSeries.Append(dataPoint);
            }

            // 分类
            pieChartSeries.RemoveAllChildren<CategoryAxisData>();
            if (model.Categories != null && model.Categories.Count() > 0)
            {
                pieChartSeries.Append(new CategoryAxisData().ChangeData(model.Categories));
            }

            // 数据
            pieChartSeries.RemoveAllChildren<Values>();
            if (model.Data != null && model.Data.Count() > 0)
            {
                pieChartSeries.Append(ChartUtil.GenerateValues(model.Data.ToArray(), null));
            }
        }
    }
}

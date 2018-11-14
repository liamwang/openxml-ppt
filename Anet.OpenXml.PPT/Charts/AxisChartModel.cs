using System.Collections.Generic;

namespace Anet.OpenXml.PPT.Charts
{
    public class AxisChartModel
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 坐标轴最大值
        /// </summary>
        public double? MaxAxisValue { get; set; }

        /// <summary>
        /// 坐标轴步长单位
        /// </summary>
        public double? MajorUnit { get; set; }

        /// <summary>
        /// 坐标轴分类
        /// </summary>
        public IEnumerable<string> Categories { get; set; }

        /// <summary>
        /// 系列数据
        /// </summary>
        public List<ChartSeriesModel> SeriesList { get; set; }
    }

    public class ChartSeriesModel
    {
        public ChartType ChartType { get; set; } = ChartType.BarChart;
        public string Title { get; set; }
        public string FormatCode { get; set; }
        public double[] Data { get; set; }
    }

    public enum ChartType : byte
    {
        BarChart = 1,
        LineChart = 2
    }
}

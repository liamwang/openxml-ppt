namespace OpenXml.PPT.Charts;

public class PieChartModel
{
    /// <summary>
    /// 标题
    /// </summary>
    public string Title { get; set; }

    /// <summary>
    /// 数据
    /// </summary>
    public IEnumerable<double> Data { get; set; }

    /// <summary>
    /// 坐标轴分类
    /// </summary>
    public IEnumerable<string> Categories { get; set; }
}

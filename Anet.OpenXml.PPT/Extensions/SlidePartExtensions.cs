using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using D = DocumentFormat.OpenXml.Drawing;

namespace Anet.OpenXml.PPT
{
    public static class SlidePartExtensions
    {
        /// <summary>
        /// 替换PPT中某一页的文字
        /// </summary>
        /// <param name="map">替换关系表</param>
        public static void ReplaceTexts(this SlidePart slidePart, Dictionary<string, string> map)
        {
            // 替换普通文本
            DoReplaceTexts(map, slidePart.Slide.Descendants<D.Text>());

            // 替换图表文本
            foreach (var diagram in slidePart.DiagramDataParts)
            {
                DoReplaceTexts(map, diagram.DataModelRoot.Descendants<D.Text>());
            }

            //// 替换图表文本
            //foreach (var chart in slidePart.ChartParts)
            //{
            //    DoReplace(map, chart.ChartSpace.Descendants<D.Text>());
            //}
        }

        private static void DoReplaceTexts(Dictionary<string, string> map, IEnumerable<D.Text> texts)
        {
            foreach (D.Text t in texts)
            {
                foreach (var kv in map)
                {
                    t.Text = t.Text.Replace(kv.Key, kv.Value ?? "");
                }
            }
        }
    }
}

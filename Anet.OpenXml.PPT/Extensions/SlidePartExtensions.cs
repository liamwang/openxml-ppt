using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace Anet.OpenXml.PPT
{
    public static class SlidePartExtensions
    {
        public static void InsertImage(this SlidePart slidePart, string base64String, long x, long y, long Cx, long Cy)
        {
            var embedId = slidePart.NewPartId();
            var nonVisualPictureProperties = new NonVisualPictureProperties(
                new NonVisualDrawingProperties() { Id = 4U, Name = "Picture" },
                new NonVisualPictureDrawingProperties(new D.PictureLocks() { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties());

            var picture = new Picture();
            var blipFill = new BlipFill();
            var blip = new D.Blip() { Embed = embedId };

            var blipExtensionList = new D.BlipExtensionList();
            var blipExtension = new D.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            var useLocalDpi = new A14.UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");


            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            blip.Append(blipExtensionList);

            var stretch = new D.Stretch();
            var fillRectangle = new D.FillRectangle();
            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);

            var shapeProperties = new ShapeProperties();

            var transform2D = new D.Transform2D();
            var offset = new D.Offset() { X = x, Y = y };
            var extents = new D.Extents() { Cx = Cx, Cy = Cy };

            transform2D.Append(offset);
            transform2D.Append(extents);

            var presetGeometry = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
            var adjustValueList = new D.AdjustValueList();

            presetGeometry.Append(adjustValueList);

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(picture);

            ImagePart imagePart = slidePart.AddNewPart<ImagePart>("image/png", embedId);
            using (var stream = new MemoryStream(Convert.FromBase64String(base64String)))
            {
                imagePart.FeedData(stream);
            }
        }

        public static string NewPartId(this SlidePart slidePart)
        {
            var idList = slidePart.Parts
                .Where(x => x.RelationshipId.StartsWith("rId"));

            if (idList.Count() == 0)
                return "rId100";

            var maxId = idList.Max(x => x.RelationshipId)
                 .Replace("rId", "");

            return "rId" + (int.Parse(maxId) + 1);
        }

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

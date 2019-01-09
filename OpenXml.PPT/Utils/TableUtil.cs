using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using D = DocumentFormat.OpenXml.Drawing;

namespace OpenXml.PPT
{
    public class TableUtil
    {
        public static D.Table CreateTable(SlidePart slidePart, List<string[]> dataSource, long x, long y, uint id = 1)
        {
            var table = NewTable(dataSource);

            var graphicFrame = new GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties() { Id = id, Name = "" },
                    new NonVisualGraphicFrameDrawingProperties(new D.GraphicFrameLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(
                    new D.Offset() { X = x, Y = y },
                    new D.Extents() { Cx = 10220195L, Cy = 4409440L }
                ),
                new D.Graphic(new D.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" })
            );
            slidePart.Slide.CommonSlideData.ShapeTree.Append(graphicFrame);

            return table;
        }

        private static D.Table NewTable(List<string[]> dataSource)
        {
            var table = new D.Table();

            var tableProperties = new D.TableProperties() { }; 
            tableProperties.Append(new D.TableStyleId() { Text = "{2A488322-F2BA-4B5B-9748-0D474271808F}" });

            var tableGrid = new D.TableGrid();
            for (var i = 0; i < dataSource[0].Length; i++)
            {
                tableGrid.Append(new D.GridColumn() { Width = 864000L });
            }

            table.Append(tableProperties);
            table.Append(tableGrid);

            for (int row = 0; row < dataSource.Count; row++)
            {
                var tableRow = new D.TableRow() { Height = 280670L };
                for (int col = 0; col < dataSource[0].Length; col++)
                {
                    tableRow.Append(NewCell(dataSource[row][col]));
                }
                table.Append(tableRow);
            }

            return table;
        }

        private static D.TableCell NewCell(string text, int fontSize = 1100, D.TextAlignmentTypeValues alignment = D.TextAlignmentTypeValues.Center)
        {
            // a:tc(TableCell)->a:txbody(TextBody)->a:p(Paragraph)->a:r(Run)->a:t(Text)
            var tableCell = new D.TableCell();
            var textBody = new D.TextBody();
            var paragraph = new D.Paragraph(new D.ParagraphProperties() { Alignment = alignment, FontAlignment = D.TextFontAlignmentValues.Center });
            var run = new D.Run();

            run.Append(new D.RunProperties { Language = "zh-CN", Dirty = false, FontSize = fontSize });
            run.Append(new D.Text() { Text = text });

            paragraph.Append(run);
            paragraph.Append(new D.EndParagraphRunProperties() { Language = "zh-CN", Dirty = false });

            textBody.Append(new D.BodyProperties());
            textBody.Append(new D.ListStyle());
            textBody.Append(paragraph);

            tableCell.Append(textBody);
            tableCell.Append(new D.TableCellProperties());

            return tableCell;
        }
    }
}

using DocumentFormat.OpenXml.Presentation;

namespace OpenXml.PPT.Defines;

public class CommonSlideDefine
{
    public static CommonSlideData NewBlankCommonSlideData() => new(
        new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new GroupShapeProperties()
        )
    );
}

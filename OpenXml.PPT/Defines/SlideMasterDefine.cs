using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXml.PPT.Defines;

public class SlideMasterDefine
{
    public static SlideMaster NewBlankSlideMaster(string slideLayoutId = "rId1") => new(
        CommonSlideDefine.NewBlankCommonSlideData(),
        new P.ColorMap()
        {
            Text1 = ColorSchemeIndexValues.Dark1,
            Background1 = ColorSchemeIndexValues.Light1,
            Text2 = ColorSchemeIndexValues.Dark2,
            Background2 = ColorSchemeIndexValues.Light2,
            Accent1 = ColorSchemeIndexValues.Accent1,
            Accent2 = ColorSchemeIndexValues.Accent2,
            Accent3 = ColorSchemeIndexValues.Accent3,
            Accent4 = ColorSchemeIndexValues.Accent4,
            Accent5 = ColorSchemeIndexValues.Accent5,
            Accent6 = ColorSchemeIndexValues.Accent6,
            Hyperlink = ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink
        },
        new SlideLayoutIdList(new SlideLayoutId() { Id = 2147483649U, RelationshipId = slideLayoutId })
    );
}

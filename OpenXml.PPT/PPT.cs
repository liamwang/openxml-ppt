using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OpenXml.PPT.Defines;

namespace OpenXml.PPT;

public class PPT : IDisposable
{
    public static MemoryStream OpenAsStream(string path)
    {
        var bytes = File.ReadAllBytes(path);
        return new MemoryStream(bytes) { Position = 0 };
    }

    public PPT()
    {
        Stream = new MemoryStream();
        Document = PresentationDocument.Create(Stream, PresentationDocumentType.Presentation);
        InitDocument(Document);
    }

    public MemoryStream Stream { get; }
    public PresentationDocument Document { get; }

    private static void InitDocument(PresentationDocument document)
    {
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        presentationPart.Presentation.Append(
            new SlideMasterIdList(new SlideMasterId() { Id = 2147483648U, RelationshipId = "rId1" }),
            new SlideIdList(new SlideId() { Id = 256U, RelationshipId = "rId2" }),
            new SlideSize() { Cx = 12192000, Cy = 6858000 },
            new NotesSize() { Cx = 6858000, Cy = 9144000 },
            new DefaultTextStyle()
        );

        var slidePart = document.CreateSlidePart("rId2"); // presentationPart.AddNewPart<SlidePart>("rId2");

        // Create SlideLayoutPart
        var slideLayoutPart = slidePart.AddNewPart<SlideLayoutPart>("rId1");
        slideLayoutPart.SlideLayout = new SlideLayout(CommonSlideDefine.NewBlankCommonSlideData());

        // Create SlideMasterPart
        var slideMasterPart = slideLayoutPart.AddNewPart<SlideMasterPart>("rId1");
        slideMasterPart.SlideMaster = SlideMasterDefine.NewBlankSlideMaster("rId1");


        // Create Theme
        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId5");
        themePart.Theme = ThemeDefine.NewBlankTheme();

        slideMasterPart.AddPart(slideLayoutPart, "rId1"); 
        presentationPart.AddPart(slideMasterPart, "rId1");  // Id 必须属于 SlideMasterIdList
        presentationPart.AddPart(themePart, "rId5"); 

        // Remove SlidePart
        //presentationPart.Presentation.SlideIdList.RemoveAllChildren();
        //presentationPart.DeletePart(slidePart);
    }

    public void Dispose()
    {
        Document?.Dispose();
        Stream?.Dispose();
        GC.SuppressFinalize(this);
    }
}

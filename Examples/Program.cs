using OpenXml.PPT;


using var ppt = new PPT();

var slidePart = ppt.Document.GetSlide(1);

var tableData = new List<string[]> {
    new string[]{"Head1","Head2"},
    new string[]{"Cell1","Cell2"},
};

TableUtil.CreateTable(slidePart, tableData, 0, 0);

ppt.Document.SaveAs(@"D:\\test.pptx");

Console.WriteLine("Done!");
Console.ReadLine();

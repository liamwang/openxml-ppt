using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXml.PPT.Charts;

public static class CategoryAxisDataExtensions
{
    public static CategoryAxisData ChangeData(
        this CategoryAxisData categoryAxisData, 
        IEnumerable<string> data)
    {
        // 先清空
        categoryAxisData.RemoveAllChildren();

        // 再拼装
        var stringCache = new StringCache();
        stringCache.Append(new PointCount { Val = (uint)data.Count() });

        for (var i = 0; i < data.Count(); i++)
        {
            var stringPoint = new StringPoint { Index = (uint)i };
            var numericValue = new NumericValue { Text = data.ElementAt(i)};
            stringPoint.Append(numericValue);
            stringCache.Append(stringPoint);
        }

        var stringReference = new StringReference();
        stringReference.Append(stringCache);

        categoryAxisData.Append(stringReference);

        return categoryAxisData;
    }
}

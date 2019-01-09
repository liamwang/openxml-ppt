using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;

namespace OpenXml.PPT.Charts
{
    public class ChartUtil
    {
        public static Values GenerateValues(double[] data, string formatCode = null)
        {
            var values = new Values();
            var numberReference = new NumberReference();
            //numberReference.Append(new Formula { Text = "" });

            var numberingCache = new NumberingCache();
            numberingCache.Append(new FormatCode { Text = string.IsNullOrEmpty(formatCode) ? "General" : formatCode });
            numberingCache.Append(new PointCount { Val = (uint)data.Length });
            for (int i = 0; i < data.Length; i++)
            {
                var numericPoint = new NumericPoint { Index = (uint)i };
                numericPoint.Append(new NumericValue { Text = data[i].ToString() });
                numberingCache.Append(numericPoint);
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);

            return values;
        }

        public static SchemeColorValues GetAccentColor(int index)
        {
            return (SchemeColorValues)Enum.Parse(typeof(SchemeColorValues), "Accent" + (index + 1));
        }
    }
}

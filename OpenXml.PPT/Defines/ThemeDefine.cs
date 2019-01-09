using DocumentFormat.OpenXml.Drawing;

namespace OpenXml.PPT.Defines
{
    public class ThemeDefine
    {
        /// <summary>
        /// Blank ColorScheme
        /// </summary>
        public static ColorScheme NewBlankColorScheme() => new ColorScheme(
            new Dark1Color(new SystemColor() { Val = SystemColorValues.WindowText, LastColor = "000000" }),
            new Light1Color(new SystemColor() { Val = SystemColorValues.Window, LastColor = "FFFFFF" }),
            new Dark2Color(new RgbColorModelHex() { Val = "000000" }),
            new Light2Color(new RgbColorModelHex() { Val = "FFFFFF" }),
            new Accent1Color(new RgbColorModelHex() { Val = "000000" }),
            new Accent2Color(new RgbColorModelHex() { Val = "000000" }),
            new Accent3Color(new RgbColorModelHex() { Val = "000000" }),
            new Accent4Color(new RgbColorModelHex() { Val = "000000" }),
            new Accent5Color(new RgbColorModelHex() { Val = "000000" }),
            new Accent6Color(new RgbColorModelHex() { Val = "000000" }),
            new Hyperlink(new RgbColorModelHex() { Val = "000000" }),
            new FollowedHyperlinkColor(new RgbColorModelHex() { Val = "000000" })
        )
        { Name = "Blank" };


        /// <summary>
        /// Blank FontScheme
        /// </summary>
        public static FontScheme NewBlankFontScheme() => new FontScheme(
            new MajorFont(new LatinFont() { Typeface = "" }, new EastAsianFont() { Typeface = "" }, new ComplexScriptFont() { Typeface = "" }),
            new MinorFont(new LatinFont() { Typeface = "" }, new EastAsianFont() { Typeface = "" }, new ComplexScriptFont() { Typeface = "" })
        )
        { Name = "Blank" };

        /// <summary>
        /// Blank FormatScheme
        /// </summary>
        public static FormatScheme NewBlankFormatScheme() => new FormatScheme(
            new FillStyleList(new NoFill(), new NoFill(), new NoFill(), new NoFill()),
            new LineStyleList(new Outline(new NoFill()), new Outline(new NoFill()), new Outline(new NoFill())),
            new EffectStyleList(new EffectStyle(new EffectList()), new EffectStyle(new EffectList()), new EffectStyle(new EffectList())),
            new BackgroundFillStyleList(new NoFill(), new NoFill(), new NoFill(), new NoFill())
        )
        { Name = "Blank" };

        /// <summary>
        /// Blank ThemeElements
        /// </summary>
        public static ThemeElements NewBlankThemeElements() => new ThemeElements(NewBlankColorScheme(), NewBlankFontScheme(), NewBlankFormatScheme());

        /// <summary>
        /// Blank Theme
        /// </summary>
        public static Theme NewBlankTheme() => new Theme(NewBlankThemeElements()) { Name = "Blank" };
    }
}

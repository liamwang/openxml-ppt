namespace OpenXml.PPT;

/// <summary>
/// English Metric Units
/// OpenXML 的计量单位
/// </summary>
public class EMU
{
    //public static long FromPt(int pt)
    //{
    //    return pt * 12700;
    //}

    public static long FromPx(int px)
    {
        return px * 9525;
    }

    public static long FromPctX(double pct)
    {
        return (long)(12192000 * pct / 100.0);
    }

    public static long FromPctY(double pct)
    {
        return (long)(6858000 * pct / 100.0);
    }

    //public static long FromCm(double cm)
    //{
    //    return FromPt((int)(cm / 0.0353));
    //}
}

namespace QuickXL;

public readonly struct XLColor(byte r, byte g, byte b, byte? a = null)
{
    public byte R { get; } = r;
    public byte G { get; } = g;
    public byte B { get; } = b;
    public byte? A { get; } = a;

    public static XLColor Make(byte r, byte g, byte b, byte? a = null)
        => new(r, g, b, a);

    public override string ToString()
        => A.HasValue
           ? $"#{A.Value:X2}{R:X2}{G:X2}{B:X2}"
           : $"#{R:X2}{G:X2}{B:X2}";

    public static implicit operator string(XLColor color)
        => color.ToString();
}

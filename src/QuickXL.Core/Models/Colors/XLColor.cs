using NPOI.XSSF.UserModel;
using System.Drawing;

namespace QuickXL.Core.Models.Colors
{
    public readonly struct XLColor(byte r, byte g, byte b, byte? a = null)
    {
        public int R { get; } = r;
        public int G { get; } = g;
        public int B { get; } = b;
        public int? A { get; } = a;


        public static implicit operator XSSFColor(XLColor color)
        {
            return color.A.HasValue
                ? new XSSFColor(Color.FromArgb(color.A.Value, color.R, color.G, color.B))
                : new XSSFColor(Color.FromArgb(color.R, color.G, color.B));
        }
    }
}

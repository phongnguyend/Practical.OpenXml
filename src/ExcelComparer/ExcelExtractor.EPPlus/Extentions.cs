using ExcelExtractor.Abstractions;
using OfficeOpenXml.Style;

namespace ExcelExtractor.EPPlus;

public static class Extentions
{
    public static Color GetColor(this ExcelColor color)
    {
        var colorModel = new Color
        {
            Theme = color.Theme.ToString().NullIfEmpty(),
            Indexed = color.Indexed < 0 ? null : color.Indexed.NullIfZero(),
            Rgb = color.Rgb.NullIfEmpty(),
            Tint = color.Tint.NullIfZero()
        };

        return colorModel;
    }
}

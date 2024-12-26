using ExcelComparer.Abstractions;
using OfficeOpenXml.Style;

namespace ExcelExtractor.EPPlus;

public static class Extentions
{
    public static Color GetColor(this ExcelColor color)
    {
        var colorModel = new Color
        {
            Theme = color.Theme.ToString(),
            Indexed = color.Indexed < 0 ? null : color.Indexed,
            Rgb = color.Rgb,
            Tint = color.Tint
        };

        return colorModel;
    }
}

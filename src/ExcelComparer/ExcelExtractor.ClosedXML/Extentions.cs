using ClosedXML.Excel;
using ExcelExtractor.Abstractions;

namespace ExcelExtractor.ClosedXML;

public static class Extentions
{
    public static Color GetColor(this XLColor color)
    {
        var colorModel = new Color();

        switch (color.ColorType)
        {
            case XLColorType.Color:
                colorModel.Rgb = color.ToString();
                break;
            case XLColorType.Theme:
                colorModel.Theme = color.ThemeColor.ToString();
                break;
            case XLColorType.Indexed:
                colorModel.Indexed = color.Indexed;
                break;
            default:
                break;
        }

        return colorModel;
    }
}

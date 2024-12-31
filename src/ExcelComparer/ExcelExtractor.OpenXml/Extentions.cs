using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExtractor.OpenXml;

public static class Extentions
{
    public static string GetText(this Cell cell, SpreadsheetDocument document)
    {
        if (cell == null)
        {
            return null;
        }

        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            SharedStringTable sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
            return sharedStringTable.ElementAt(int.Parse(cell.CellValue?.Text)).InnerText;
        }
        else
        {
            return cell.CellValue?.Text;
        }
    }

    public static Abstractions.Style GetStyle(this Cell cell, SpreadsheetDocument document)
    {
        if (cell == null)
        {
            return null;
        }

        Stylesheet stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
        var style = (CellFormat)stylesheet.CellFormats.ElementAt(int.Parse(cell.StyleIndex));

        var font = (Font)stylesheet.Fonts.ElementAt(int.Parse(style.FontId));

        var fill = (Fill)stylesheet.Fills.ElementAt(int.Parse(style.FillId));

        var border = (Border)stylesheet.Borders.ElementAt(int.Parse(style.BorderId));

        var format = (CellFormat)stylesheet.CellFormats.ElementAt(int.Parse(style.FormatId));

        var numberingFormat = style.ApplyNumberFormat != null ? (NumberingFormat)stylesheet.NumberingFormats?.ElementAt(int.Parse(format.NumberFormatId)) : null;

        var rs = new Abstractions.Style
        {
            Font = new Abstractions.Font
            {
                Color = GetColor(font.Color, document),
                Size = font.FontSize.Val,
                Bold = font.Bold != null
            },
            Fill = new Abstractions.Fill
            {
                PatternType = Enum.Parse(typeof(FillPatternType), fill.PatternFill.PatternType, ignoreCase: true).ToString(),
                BackgroundColor = fill.PatternFill.BackgroundColor.GetColor(document)
            },
            Alignment = new Abstractions.Alignment
            {
                Horizontal = Enum.Parse(typeof(AlignmentHorizontal), style.Alignment?.Horizontal?.InnerText ?? "General", ignoreCase: true).ToString(),
                Vertical = Enum.Parse(typeof(AlignmentVertical), style.Alignment?.Vertical?.InnerText ?? "Bottom", ignoreCase: true).ToString()
            },
            Border = new Abstractions.Border
            {
                Top = new Abstractions.BorderItem
                {
                    Color = border.TopBorder.Color.GetColor(document),
                    Style = FormatBorderStyle(border.TopBorder.Style.ToString())
                },
                Left = new Abstractions.BorderItem
                {
                    Color = border.LeftBorder.Color.GetColor(document),
                    Style = FormatBorderStyle(border.LeftBorder.Style.ToString())
                },
                Bottom = new Abstractions.BorderItem
                {
                    Color = border.BottomBorder.Color.GetColor(document),
                    Style = FormatBorderStyle(border.BottomBorder.Style.ToString())
                },
                Right = new Abstractions.BorderItem
                {
                    Color = border.RightBorder.Color.GetColor(document),
                    Style = FormatBorderStyle(border.RightBorder.Style.ToString())
                }
            },
            Numberformat = new Abstractions.Numberformat
            {
                Format = numberingFormat?.FormatCode
            }
        };

        return rs;
    }


    public static Abstractions.Color GetColor(this Color color, SpreadsheetDocument document)
    {
        if (color == null)
        {
            return null;
        }

        var colorModel = new Abstractions.Color();

        if (color.Theme != null)
        {
            colorModel.Theme = ((ThemeColor)(int)color.Theme.Value).ToString();
        }

        colorModel.Tint = color.Tint != null ? (decimal)color.Tint.Value : null;
        colorModel.Indexed = color.Indexed != null ? (int)color.Indexed.Value : null;
        colorModel.Rgb = color.Rgb?.Value;


        return colorModel;
    }

    public static Abstractions.Color GetColor(this BackgroundColor color, SpreadsheetDocument document)
    {

        if (color == null)
        {
            return null;
        }

        var colorModel = new Abstractions.Color();

        if (color.Theme != null)
        {
            colorModel.Theme = ((ThemeColor)(int)color.Theme.Value).ToString();
        }

        colorModel.Tint = color.Tint != null ? (decimal)color.Tint.Value : null;
        colorModel.Indexed = color.Indexed != null ? (int)color.Indexed.Value : null;
        colorModel.Rgb = color.Rgb?.Value;


        return colorModel;
    }

    private static string FormatBorderStyle(this string style)
    {
        return Enum.Parse(typeof(BorderStyle), style, ignoreCase: true).ToString();
    }
}

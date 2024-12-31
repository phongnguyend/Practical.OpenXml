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
                PatternType = fill.PatternFill.PatternType,
                BackgroundColor = fill.PatternFill.BackgroundColor.GetColor(document)
            },
            Alignment = new Abstractions.Alignment
            {
                Horizontal = style.Alignment?.Horizontal?.InnerText,
                Vertical = style.Alignment?.Vertical?.InnerText
            },
            //Border = new Border
            //{
            //    Top = new BorderItem
            //    {
            //        Color = cell.Style.Border.TopBorderColor.GetColor(),
            //        Style = cell.Style.Border.TopBorder.ToString()
            //    },
            //    Left = new BorderItem
            //    {
            //        Color = cell.Style.Border.LeftBorderColor.GetColor(),
            //        Style = cell.Style.Border.LeftBorder.ToString()
            //    },
            //    Bottom = new BorderItem
            //    {
            //        Color = cell.Style.Border.BottomBorderColor.GetColor(),
            //        Style = cell.Style.Border.BottomBorder.ToString()
            //    },
            //    Right = new BorderItem
            //    {
            //        Color = cell.Style.Border.RightBorderColor.GetColor(),
            //        Style = cell.Style.Border.RightBorder.ToString()
            //    }
            //},
            //Numberformat = new Numberformat
            //{
            //    Format = cell.Style.NumberFormat.Format
            //}
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

    public static Abstractions.Color GetColor(this DocumentFormat.OpenXml.Spreadsheet.BackgroundColor color, SpreadsheetDocument document)
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
}

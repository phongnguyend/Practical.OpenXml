
using ExcelExtractor.Abstractions;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExtractor.NPOI;

public class NPOIExcelExtractor : IExcelExtractor
{
    public Workbook Extract(Stream stream)
    {
        var workbook = new Workbook
        {
            Worksheets = new List<Worksheet>()
        };

        var npoiWorkbook = new XSSFWorkbook(stream);

        for (int i = 0; i < npoiWorkbook.NumberOfSheets; i++)
        {
            var sheet = npoiWorkbook.GetSheetAt(i);

            var cells = new List<Cell>();

            workbook.Worksheets.Add(new Worksheet
            {
                Name = sheet.SheetName,
                Cells = cells
            });

            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);

                        if (cell is null)
                            continue;

                        var cellStyle = cell.CellStyle as XSSFCellStyle;
                        var font = npoiWorkbook.GetFontAt(cellStyle.FontIndex) as XSSFFont;

                        var myCell = new Cell
                        {
                            Row = rowIndex + 1,
                            Column = colIndex + 1,
                            Value = cell.ToString(),
                            Merged = cell.IsMergedCell,
                            Style = new Style
                            {
                                Font = new Font
                                {
                                    Color = GetColor(font.GetXSSFColor()),
                                    Size = font.FontHeightInPoints,
                                    Bold = font.IsBold
                                },
                                Fill = new Fill
                                {
                                    PatternType = GetFillPatternType(cellStyle),
                                    BackgroundColor = GetColor(cellStyle.FillBackgroundXSSFColor)
                                },
                                Alignment = new Alignment
                                {
                                    Horizontal = cellStyle.Alignment.ToString(),
                                    Vertical = cellStyle.VerticalAlignment.ToString()
                                },
                                Border = new Border
                                {
                                    Top = new BorderItem
                                    {
                                        Color = GetColor(cellStyle.TopBorderXSSFColor),
                                        Style = cellStyle.BorderTop.ToString()
                                    },
                                    Left = new BorderItem
                                    {
                                        Color = GetColor(cellStyle.LeftBorderXSSFColor),
                                        Style = cellStyle.BorderLeft.ToString()
                                    },
                                    Bottom = new BorderItem
                                    {
                                        Color = GetColor(cellStyle.BottomBorderXSSFColor),
                                        Style = cellStyle.BorderBottom.ToString()
                                    },
                                    Right = new BorderItem
                                    {
                                        Color = GetColor(cellStyle.RightBorderXSSFColor),
                                        Style = cellStyle.BorderRight.ToString()
                                    }
                                },
                                Numberformat = new Numberformat
                                {
                                    Format = GetNumberformat(cellStyle)
                                }
                            }
                        };

                        cells.Add(myCell);
                    }
                }
            }
        }

        return workbook;
    }

    private static string GetNumberformat(XSSFCellStyle cellStyle)
    {
        var temp = cellStyle.GetDataFormatString();

        if (temp == "General")
        {
            return null;
        }

        return temp;
    }

    private static string GetFillPatternType(XSSFCellStyle cellStyle)
    {
        var temp = ((ST_PatternType)(int)cellStyle.FillPattern).ToString();

        return Enum.Parse(typeof(FillPatternType), temp, ignoreCase: true).ToString();
    }

    private static Color GetColor(XSSFColor color)
    {
        if (color is null)
            return null;

        var colorModel = new Color
        {
            Theme = color.IsThemed ? ((ThemeColor)color.Theme).ToString() : null,
            Indexed = color.IsIndexed ? color.Indexed : null,
            Rgb = color.IsRGB ? color.ARGBHex : null,
            Tint = color.HasTint ? (decimal)color.Tint : null
        };

        if(color.IsThemed)
        {
            colorModel.Rgb = null;
        }

        return colorModel;
    }
}

public enum ThemeColor
{
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink
}

public enum FillPatternType
{
    DarkDown,
    DarkGray,
    DarkGrid,
    DarkHorizontal,
    DarkTrellis,
    DarkUp,
    DarkVertical,
    Gray0625,
    Gray125,
    LightDown,
    LightGray,
    LightGrid,
    LightHorizontal,
    LightTrellis,
    LightUp,
    LightVertical,
    MediumGray,
    None,
    Solid
}

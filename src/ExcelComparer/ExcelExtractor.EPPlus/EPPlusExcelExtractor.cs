using ExcelComparer.Abstractions;
using OfficeOpenXml;

namespace ExcelExtractor.EPPlus;

public class EPPlusExcelExtractor : IExcelExtractor
{
    public Workbook Extract(Stream stream)
    {
        var workbook = new Workbook
        {
            Worksheets = new List<Worksheet>()
        };

        using var packge = new ExcelPackage(stream);

        foreach (var worksheet in packge.Workbook.Worksheets)
        {
            var cells = new List<Cell>();

            workbook.Worksheets.Add(new Worksheet
            {
                Name = worksheet.Name,
                Cells = cells
            });

            for (var i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (var j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    var cell = worksheet.Cells[i, j];

                    var myCell = new Cell
                    {
                        Row = i,
                        Column = j,
                        Value = cell.GetValue<string>(),
                        Merged = cell.Merge,
                        Style = new Style
                        {
                            Font = new Font
                            {
                                Color = cell.Style.Font.Color.GetColor(),
                                Size = cell.Style.Font.Size,
                                Bold = cell.Style.Font.Bold
                            },
                            Fill = new Fill
                            {
                                PatternType = cell.Style.Fill.PatternType.ToString(),
                                BackgroundColor = cell.Style.Fill.BackgroundColor.GetColor()
                            },
                            Alignment = new Alignment
                            {
                                Horizontal = cell.Style.HorizontalAlignment.ToString(),
                                Vertical = cell.Style.VerticalAlignment.ToString()
                            },
                            Border = new Border
                            {
                                Top = new BorderItem
                                {
                                    Color = cell.Style.Border.Top.Color.GetColor(),
                                    Style = cell.Style.Border.Top.Style.ToString()
                                },
                                Left = new BorderItem
                                {
                                    Color = cell.Style.Border.Left.Color.GetColor(),
                                    Style = cell.Style.Border.Left.Style.ToString()
                                },
                                Bottom = new BorderItem
                                {
                                    Color = cell.Style.Border.Bottom.Color.GetColor(),
                                    Style = cell.Style.Border.Bottom.Style.ToString()
                                },
                                Right = new BorderItem
                                {
                                    Color = cell.Style.Border.Right.Color.GetColor(),
                                    Style = cell.Style.Border.Right.Style.ToString()
                                }
                            },
                            Numberformat = new Numberformat
                            {
                                Format = cell.Style.Numberformat.Format
                            }
                        }
                    };

                    cells.Add(myCell);
                }
            }
        }

        return workbook;
    }
}

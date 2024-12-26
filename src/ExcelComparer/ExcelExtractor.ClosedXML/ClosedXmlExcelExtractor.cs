using ClosedXML.Excel;
using ExcelComparer.Abstractions;

namespace ExcelExtractor.ClosedXML;

public class ClosedXmlExcelExtractor : IExcelExtractor
{
    public Workbook Extract(Stream stream)
    {
        var workbook = new Workbook
        {
            Worksheets = new List<Worksheet>()
        };

        using var closedXmlWorkbook = new XLWorkbook(stream);
        foreach (var worksheet in closedXmlWorkbook.Worksheets)
        {
            var cells = new List<Cell>();

            workbook.Worksheets.Add(new Worksheet
            {
                Name = worksheet.Name,
                Cells = cells
            });

            for (var i = 1; i <= worksheet.LastRowUsed().RowNumber(); i++)
            {
                for (var j = 1; j <= worksheet.LastColumnUsed().ColumnNumber(); j++)
                {
                    var cell = worksheet.Cell(i, j);

                    var myCell = new Cell
                    {
                        Row = i,
                        Column = j,
                        Value = cell.GetString(),
                        Merged = cell.IsMerged(),
                        Style = new Style
                        {
                            Font = new Font
                            {
                                Color = cell.Style.Font.FontColor.GetColor(),
                                Size = cell.Style.Font.FontSize,
                                Bold = cell.Style.Font.Bold
                            },
                            Fill = new Fill
                            {
                                PatternType = cell.Style.Fill.PatternType.ToString(),
                                BackgroundColor = cell.Style.Fill.BackgroundColor.GetColor()
                            },
                            Alignment = new Alignment
                            {
                                Horizontal = cell.Style.Alignment.Horizontal.ToString(),
                                Vertical = cell.Style.Alignment.Vertical.ToString()
                            },
                            Border = new Border
                            {
                                Top = new BorderItem
                                {
                                    Color = cell.Style.Border.TopBorderColor.GetColor(),
                                    Style = cell.Style.Border.TopBorder.ToString()
                                },
                                Left = new BorderItem
                                {
                                    Color = cell.Style.Border.LeftBorderColor.GetColor(),
                                    Style = cell.Style.Border.LeftBorder.ToString()
                                },
                                Bottom = new BorderItem
                                {
                                    Color = cell.Style.Border.BottomBorderColor.GetColor(),
                                    Style = cell.Style.Border.BottomBorder.ToString()
                                },
                                Right = new BorderItem
                                {
                                    Color = cell.Style.Border.RightBorderColor.GetColor(),
                                    Style = cell.Style.Border.RightBorder.ToString()
                                }
                            },
                            Numberformat = new Numberformat
                            {
                                Format = cell.Style.NumberFormat.Format
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

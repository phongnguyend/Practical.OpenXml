using ExcelDataReader;
using ExcelExtractor.Abstractions;

namespace ExcelExtractor.ExcelDataReader;

public class ExcelDataReaderExtractor : IExcelExtractor
{
    public Workbook Extract(Stream stream)
    {
        var workbook = new Workbook
        {
            Worksheets = new List<Worksheet>()
        };

        var reader = ExcelReaderFactory.CreateReader(stream);

        do
        {
            //if (reader.VisibleState != VisibleStateConstants.Visible)
            //{
            //    continue;
            //}

            var sheetName = reader.Name;

            var cells = new List<Cell>();

            workbook.Worksheets.Add(new Worksheet
            {
                Name = sheetName,
                Cells = cells
            });

            var mergeCells = reader.MergeCells.Select(x => new MergeCell
            {
                FromRow = x.FromRow + 1,
                FromColumn = x.FromColumn + 1,
                ToRow = x.ToRow + 1,
                ToColumn = x.ToColumn + 1
            }).ToList();

            int rowIndex = 0;

            while (reader.Read())
            {
                for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                {
                    var style = reader.GetCellStyle(columnIndex);

                    var myCell = new Cell
                    {
                        Row = rowIndex + 1,
                        Column = columnIndex + 1,
                        Value = reader.GetValue(columnIndex)?.ToString(),
                        Merged = mergeCells.Any(x => x.Contains(rowIndex + 1, columnIndex + 1)),
                        Style = new Style
                        {
                            //Font = new Font
                            //{
                            //    Color = cell.Style.Font.FontColor.GetColor(),
                            //    Size = cell.Style.Font.FontSize,
                            //    Bold = cell.Style.Font.Bold
                            //},
                            //Fill = new Fill
                            //{
                            //    PatternType = cell.Style.Fill.PatternType.ToString(),
                            //    BackgroundColor = cell.Style.Fill.BackgroundColor.GetColor()
                            //},
                            Alignment = new Alignment
                            {
                                Horizontal = style.HorizontalAlignment.ToString(),
                                //Vertical = style.VerticalAlignment.ToString(),
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
                            Numberformat = new Numberformat
                            {
                                Format = reader.GetNumberFormatString(columnIndex)
                            }
                        }
                    };

                    cells.Add(myCell);
                }

                rowIndex++;
            }

        } while (reader.NextResult());

        return workbook;
    }
}

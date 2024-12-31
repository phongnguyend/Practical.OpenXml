using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExtractor.Abstractions;
using OpenXmlRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using Cell = ExcelExtractor.Abstractions.Cell;
using Workbook = ExcelExtractor.Abstractions.Workbook;
using Worksheet = ExcelExtractor.Abstractions.Worksheet;

namespace ExcelExtractor.OpenXml;

public class OpenXmlExtractor : IExcelExtractor
{
    public Workbook Extract(Stream stream)
    {
        var workbook = new Workbook
        {
            Worksheets = new List<Worksheet>()
        };

        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        SharedStringTable sharedStringTable = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;

        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

        foreach (Sheet sheet in workbookPart.Workbook.Sheets)
        {

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            var cells = new List<Cell>();

            workbook.Worksheets.Add(new Worksheet
            {
                Name = sheet.Name,
                Cells = cells
            });

            int i = 1;

            foreach (OpenXmlRow r in sheetData.Elements<OpenXmlRow>())
            {
                int j = 1;

                foreach (OpenXmlCell cell in r.Elements<OpenXmlCell>())
                {
                    (var row, var column) = ExcelHelper.ConvertAddressToIndex(cell.CellReference);
                    var myCell = new Cell
                    {
                        Row = row,
                        Column = column,
                        Value = cell.GetText(spreadsheetDocument),
                        //Merged = cell.IsMerged(),
                        //Style = new Style
                        //{
                        //    Font = new Font
                        //    {
                        //        Color = cell.Style.Font.FontColor.GetColor(),
                        //        Size = cell.Style.Font.FontSize,
                        //        Bold = cell.Style.Font.Bold
                        //    },
                        //    Fill = new Fill
                        //    {
                        //        PatternType = cell.Style.Fill.PatternType.ToString(),
                        //        BackgroundColor = cell.Style.Fill.BackgroundColor.GetColor()
                        //    },
                        //    Alignment = new Alignment
                        //    {
                        //        Horizontal = cell.Style.Alignment.Horizontal.ToString(),
                        //        Vertical = cell.Style.Alignment.Vertical.ToString()
                        //    },
                        //    Border = new Border
                        //    {
                        //        Top = new BorderItem
                        //        {
                        //            Color = cell.Style.Border.TopBorderColor.GetColor(),
                        //            Style = cell.Style.Border.TopBorder.ToString()
                        //        },
                        //        Left = new BorderItem
                        //        {
                        //            Color = cell.Style.Border.LeftBorderColor.GetColor(),
                        //            Style = cell.Style.Border.LeftBorder.ToString()
                        //        },
                        //        Bottom = new BorderItem
                        //        {
                        //            Color = cell.Style.Border.BottomBorderColor.GetColor(),
                        //            Style = cell.Style.Border.BottomBorder.ToString()
                        //        },
                        //        Right = new BorderItem
                        //        {
                        //            Color = cell.Style.Border.RightBorderColor.GetColor(),
                        //            Style = cell.Style.Border.RightBorder.ToString()
                        //        }
                        //    },
                        //    Numberformat = new Numberformat
                        //    {
                        //        Format = cell.Style.NumberFormat.Format
                        //    }
                        //}
                    };

                    cells.Add(myCell);

                    j++;
                }

                i++;
            }
        }

        return workbook;
    }
}

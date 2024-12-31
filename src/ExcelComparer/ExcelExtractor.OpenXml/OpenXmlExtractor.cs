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
                        Style = cell.GetStyle(spreadsheetDocument)
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

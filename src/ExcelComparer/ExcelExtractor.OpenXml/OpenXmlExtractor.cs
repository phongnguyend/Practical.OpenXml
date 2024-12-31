using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExtractor.Abstractions;
using OpenXmlRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using OpenXmlMergeCells = DocumentFormat.OpenXml.Spreadsheet.MergeCells;
using OpenXmlMergeCell = DocumentFormat.OpenXml.Spreadsheet.MergeCell;
using Cell = ExcelExtractor.Abstractions.Cell;
using MergeCell = ExcelExtractor.Abstractions.MergeCell;
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

            var mergedCells = GetMergeCells(worksheetPart);

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
                    (var row, var column) = ExcelHelper.ParseAddress(cell.CellReference);
                    var myCell = new Cell
                    {
                        Row = row,
                        Column = column,
                        Value = cell.GetText(spreadsheetDocument),
                        Merged = mergedCells.Any(x => x.Contains(row, column)),
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

    private static List<MergeCell> GetMergeCells(WorksheetPart worksheetPart)
    {
        var mergedCells = new List<MergeCell>();

        var openXmlMergedCells = worksheetPart.Worksheet.Elements<OpenXmlMergeCells>().FirstOrDefault();

        if (openXmlMergedCells == null)
        {
            return mergedCells;
        }

        foreach (var item in openXmlMergedCells)
        {
            var xx = ExcelHelper.ParseRange(((OpenXmlMergeCell)item).Reference);

            var mergeCell = new MergeCell
            {
                FromRow = xx.FromRow,
                FromColumn = xx.FromColumn,
                ToRow = xx.ToRow,
                ToColumn = xx.ToColumn
            };

            mergedCells.Add(mergeCell);
        }

        return mergedCells;
    }
}

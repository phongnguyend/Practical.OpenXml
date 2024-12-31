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
}

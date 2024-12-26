using ExcelExtractor.ClosedXML;


var fileName = "C:\\Users\\Phong.NguyenDoan\\Downloads\\Temp\\xxx.xlsx";

using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
var extractor = new ClosedXmlExcelExtractor();
var workbook = extractor.Extract(fileStream);

foreach (var worksheet in workbook.Worksheets)
{
    Console.WriteLine($"Worksheet: {worksheet.Name}");
    foreach (var cell in worksheet.Cells)
    {
        Console.WriteLine($"[{cell.Row}, {cell.Column}]: {cell.Value}");
    }
}

File.WriteAllText("output.json", workbook.ToJson());
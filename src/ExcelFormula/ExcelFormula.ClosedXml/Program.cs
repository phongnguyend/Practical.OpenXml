using ClosedXML.Excel;

var workbook = new XLWorkbook($"../../../../Template.xlsx");
var worksheet = workbook.Worksheet("Sheet1");

//var sum = worksheet.Evaluate("SUM(A1,B1)");

worksheet.Cell("A1").Value = 1;
worksheet.Cell("B1").Value = 5;

workbook.RecalculateAllFormulas();

var num = worksheet.Cell("B2").Value.ToString();

Console.WriteLine($"The sum of A1 and B1 is: {num}");
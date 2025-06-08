using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var excelPackage = new ExcelPackage($"../../../../Template.xlsx");
var worksheet = excelPackage.Workbook.Worksheets["Sheet1"];

worksheet.Cells["A1"].Value = 1;
worksheet.Cells["B1"].Value = 5;

worksheet.Calculate();

var num = worksheet.Cells["B2"].Value.ToString();

Console.WriteLine($"The sum of A1 and B1 is: {num}");
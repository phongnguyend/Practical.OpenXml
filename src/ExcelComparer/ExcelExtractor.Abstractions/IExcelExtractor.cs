namespace ExcelExtractor.Abstractions;

public interface IExcelExtractor
{
    Workbook Extract(Stream stream);
}

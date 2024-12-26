namespace ExcelComparer.Abstractions;

public interface IExcelExtractor
{
    Workbook Extract(Stream stream);
}

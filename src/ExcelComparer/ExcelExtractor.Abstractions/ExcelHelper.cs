using System.Text.RegularExpressions;

namespace ExcelExtractor.Abstractions;

public static class ExcelHelper
{
    private static readonly Dictionary<string, int> _columnNames = new Dictionary<string, int>();
    private static readonly Dictionary<int, string> _columnIndexs = new Dictionary<int, string>();

    static ExcelHelper()
    {
        var range = Enumerable.Range(1, 16_384).
            Select(x => new { Index = x, Name = GetExcelColumnName(x) });
        _columnNames = range.ToDictionary(x => x.Name, x => x.Index);
        _columnIndexs = range.ToDictionary(x => x.Index, x => x.Name);
    }

    private static string GetExcelColumnName(int columnNumber)
    {
        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    public static int ConvertColumnNameToIndex(string columnName)
    {
        return _columnNames[columnName];
    }

    public static string ConvertColumnIndexToName(int columnIndex)
    {
        return _columnIndexs[columnIndex];
    }

    public static (int Row, int Column) ConvertAddressToIndex(string address)
    {
        var match = Regex.Match(address, @"(?<column>[A-Z]+)(?<row>\d+)");
        var column = ConvertColumnNameToIndex(match.Groups["column"].Value);
        var row = int.Parse(match.Groups["row"].Value);
        return (row, column);
    }
}

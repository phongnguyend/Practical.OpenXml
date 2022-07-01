namespace Practical.ExcelDataReader
{
    public static class ExcelConverter
    {
        private static readonly Dictionary<string, int> _columnNames = new Dictionary<string, int>();
        private static readonly Dictionary<int, string> _columnIndexs = new Dictionary<int, string>();

        static ExcelConverter()
        {
            var range = Enumerable.Range(1, 1000).
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

        public static int ColumnNameToIndex(string columnName)
        {
            return _columnNames[columnName] - 1;
        }

        public static string ColumnIndexToName(int columnIndex)
        {
            return _columnIndexs[columnIndex + 1];
        }
    }
}

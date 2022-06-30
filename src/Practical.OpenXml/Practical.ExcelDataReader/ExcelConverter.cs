namespace Practical.ExcelDataReader
{
    public static class ExcelConverter
    {
        private static readonly Dictionary<string, int> _columnNames = new Dictionary<string, int>();

        static ExcelConverter()
        {
            _columnNames = Enumerable.Range(1, 1000).
                Select(x => new { Index = x, Name = GetExcelColumnName(x) })
                .ToDictionary(x => x.Name, x => x.Index);
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

        public static Dictionary<string, int> ColumNameToIndexMappings()
        {
            return _columnNames;
        }

        public static int ColumNameToIndex(string columnName)
        {
            return _columnNames[columnName] - 1;
        }
    }
}

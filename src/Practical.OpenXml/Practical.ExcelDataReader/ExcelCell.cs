namespace Practical.ExcelDataReader
{
    public class ExcelCell
    {
        public string? Address { get; set; }

        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }

        public string? Value { get; set; }

        public string? SheetName { get; set; }
    }
}

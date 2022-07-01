using ExcelDataReader;

namespace Practical.ExcelDataReader
{
    public static class IExcelDataReaderExentions
    {
        public static bool IsEmptyRow(this IExcelDataReader reader)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (!string.IsNullOrWhiteSpace(reader.GetValue(i)?.ToString()))
                {
                    return false;
                }
            }
            return true;
        }
    }
}

namespace Practical.ExcelDataReader
{
    public static class StringExtentions
    {
        public static bool In(this string str, IEnumerable<string> list)
        {
            return list.Contains(str, StringComparer.OrdinalIgnoreCase);
        }

        public static bool NotIn(this string str, IEnumerable<string> list)
        {
            return !In(str, list);
        }
    }
}

namespace ExcelExtractor.Abstractions;

public static class Extensions
{
    public static string NullIfEmpty(this string value)
    {
        return string.IsNullOrEmpty(value) ? null : value;
    }

    public static int? NullIfZero(this int value)
    {
        return value == 0 ? null : value;
    }

    public static int? NullIfZero(this int? value)
    {
        return value == 0 ? null : value;
    }

    public static decimal? NullIfZero(this decimal value)
    {
        return value == 0 ? null : value;
    }

    public static decimal? NullIfZero(this decimal? value)
    {
        return value == 0 ? null : value;
    }
}

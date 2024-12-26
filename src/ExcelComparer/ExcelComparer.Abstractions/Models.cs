using System.Text.Json;

namespace ExcelComparer.Abstractions;

public class Workbook
{
    public List<Worksheet> Worksheets { get; set; }

    public string ToJson()
    {
        return JsonSerializer.Serialize(this, new JsonSerializerOptions
        {
            WriteIndented = true
        });
    }
}

public class Worksheet
{
    public string Name { get; set; }

    public List<Cell> Cells { get; set; }
}

public class Cell
{
    public int Row { get; set; }

    public int Column { get; set; }

    public string Address
    {
        get
        {
            return $"{ExcelHelper.ConvertColumnIndexToName(Column)}{Row}";
        }
    }

    public string Value { get; set; }

    public bool Merged { get; set; }

    public Style Style { get; set; }
}

public class Style
{
    public Font Font { get; set; }

    public Fill Fill { get; set; }

    public Alignment Alignment { get; set; }

    public Numberformat Numberformat { get; set; }

    public Border Border { get; set; }
}

public class Font
{
    public Color Color { get; set; }

    public double Size { get; set; }

    public bool Bold { get; set; }
}

public class Color
{
    public string Theme { get; set; }

    public decimal Tint { get; set; }

    public int? Indexed { get; set; }

    public string Rgb { get; set; }
}

public class Fill
{
    public string PatternType { get; set; }

    public Color BackgroundColor { get; set; }
}

public class Alignment
{
    public string Horizontal { get; set; }

    public string Vertical { get; set; }
}

public class Numberformat
{
    public string Format { get; set; }
}

public class Border
{
    public BorderItem Top { get; set; }

    public BorderItem Left { get; set; }

    public BorderItem Right { get; set; }

    public BorderItem Bottom { get; set; }
}

public class BorderItem
{
    public string Style { get; set; }

    public Color Color { get; set; }
}
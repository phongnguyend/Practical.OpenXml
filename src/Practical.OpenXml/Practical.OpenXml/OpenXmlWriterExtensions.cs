using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Practical.OpenXml
{
    public static class OpenXmlWriterExtensions
    {
        public static void CreateShareStringPart(this WorkbookPart workbookPart, SharedStringData sharedStringData)
        {
            if (sharedStringData.MaxIndex <= 0)
            {
                return;
            }

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            using var writer = OpenXmlWriter.Create(sharedStringPart);
            writer.WriteStartElement(new SharedStringTable());
            foreach (var item in sharedStringData)
            {
                writer.WriteStartElement(new SharedStringItem());
                writer.WriteElement(new Text(item.Key));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        public static void WriteInlineStringCellValue(this OpenXmlWriter writer, string cellValue, List<OpenXmlAttribute> attributes = null)
        {
            if (attributes == null)
            {
                attributes = new List<OpenXmlAttribute>();
            }

            attributes.Add(new OpenXmlAttribute("t", null, "inlineStr"));
            writer.WriteStartElement(new Cell(), attributes);
            writer.WriteElement(new InlineString(new Text(cellValue)));
            writer.WriteEndElement();
        }

        public static void WriteSharedStringCellValue(this OpenXmlWriter writer, string cellValue, SharedStringData sharedStringData, List<OpenXmlAttribute> attributes = null)
        {
            if (attributes == null)
            {
                attributes = new List<OpenXmlAttribute>();
            }

            attributes.Add(new OpenXmlAttribute("t", null, "s")); // shared string type
            writer.WriteStartElement(new Cell(), attributes);

            if (!sharedStringData.ContainsKey(cellValue))
            {
                sharedStringData.Add(cellValue, sharedStringData.MaxIndex);
                sharedStringData.MaxIndex++;
            }

            //writing the index as the cell value
            writer.WriteElement(new CellValue(sharedStringData[cellValue].ToString()));
            writer.WriteEndElement();
        }

        public static void WriteDateCellValue(this OpenXmlWriter writer, DateTime? cellValue, List<OpenXmlAttribute> attributes = null)
        {
            if (attributes == null)
            {
                attributes = new List<OpenXmlAttribute>();
            }

            writer.WriteStartElement(new Cell(), attributes);
            writer.WriteElement(cellValue.HasValue ? new CellValue(cellValue.Value.ToOADate().ToString()) : new CellValue());
            writer.WriteEndElement();
        }

        public static void WriteBooleanCellValue(this OpenXmlWriter writer, string cellValue, List<OpenXmlAttribute> attributes = null)
        {
            if (attributes == null)
            {
                attributes = new List<OpenXmlAttribute>();
            }

            attributes.Add(new OpenXmlAttribute("t", null, "b")); // boolean type
            writer.WriteStartElement(new Cell(), attributes);
            writer.WriteElement(new CellValue(cellValue == "True" ? "1" : "0"));
            writer.WriteEndElement();
        }

        public static void WriteDecimalCellValue(this OpenXmlWriter writer, decimal? cellValue, List<OpenXmlAttribute> attributes = null)
        {
            if (attributes == null)
            {
                attributes = new List<OpenXmlAttribute>();
            }

            writer.WriteStartElement(new Cell(), attributes);
            writer.WriteElement(cellValue.HasValue? new CellValue(cellValue.Value): new CellValue());
            writer.WriteEndElement();
        }
    }
}

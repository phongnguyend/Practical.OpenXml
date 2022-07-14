using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Practical.OpenXml
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = Path.Combine(path, @"test.xlsx");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var headerList = new string[] { "Header 1", "Header 2", "Header 3", "Header 4", "Decimal Column", "DateTime Column", "Rotated and Centered" };

            var stopWatch = new Stopwatch();

            using (var spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                var workbookPart = spreadSheet.AddWorkbookPart();

                SaveCustomStylesheet(workbookPart);

                var dateTimeCellFormatIndex = 1;
                var decimalCellFormatIndex = 2;
                var headerCellFormatIndex = 3;
                var rotatedAndCenteredCellFormatIndex = 4;

                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild(new Sheets());

                // create worksheet 1
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);

                var sharedStringData = new SharedStringData();

                Console.WriteLine("Starting to generate 1 million random string");

                //1 million rows
                var a = new string[1000000, 4];
                var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var random = new Random();
                for (int i = 0; i < 100; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        a[i, j] = new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
                    }
                }

                Console.WriteLine("Starting to generate 1 million excel rows...");

                stopWatch.Start();

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    //Create header row
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < headerList.Length; i++)
                    {
                        if (headerList[i] == "Rotated and Centered")
                        {
                            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{rotatedAndCenteredCellFormatIndex}") }.ToList();
                            writer.WriteSharedStringCellValue(headerList[i], sharedStringData, attributes);
                        }
                        else
                        {
                            //header formatting attribute. This will create a <c> element with s=3 as its attribute
                            //s stands for styleindex
                            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{headerCellFormatIndex}") }.ToList();
                            writer.WriteSharedStringCellValue(headerList[i], sharedStringData, attributes);
                        }

                    }
                    writer.WriteEndElement(); //end of Row tag

                    for (int i = 0; i < 100; i++)
                    {
                        writer.WriteStartElement(new Row());
                        for (int j = 0; j < 4; j++)
                        {
                            writer.WriteInlineStringCellValue(a[i, j]);
                        }

                        if (i % 5 == 0)
                        {
                            writer.WriteDecimalCellValue(1000.01m, new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{decimalCellFormatIndex}") }.ToList());
                            writer.WriteDateCellValue(null, new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{dateTimeCellFormatIndex}") }.ToList());
                        }
                        else
                        {
                            writer.WriteDecimalCellValue(null, new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{decimalCellFormatIndex}") }.ToList());
                            writer.WriteDateCellValue(DateTime.Now, new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, $"{dateTimeCellFormatIndex}") }.ToList());
                        }

                        writer.WriteEndElement(); //end of Row tag

                    }

                    writer.WriteEndElement(); //end of SheetData
                    writer.WriteEndElement(); //end of worksheet
                    writer.Close();
                }

                //create the share string part using sax like approach too
                workbookPart.CreateShareStringPart(sharedStringData);
            }

            stopWatch.Stop();

            Console.WriteLine(string.Format("Time elapsed for writing 1 million rows: {0}", stopWatch.Elapsed));
            Console.ReadLine();
        }

        private static Stylesheet CreateDefaultStylesheet()
        {
            var stylesheet = new Stylesheet();

            var fonts = new Fonts();
            var font = new Font();
            var fontName = new FontName
            {
                Val = "Calibri"
            };
            var fontSize = new FontSize
            {
                Val = 11
            };
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);
            fonts.Count = (uint)fonts.ChildElements.Count;

            font = new Font
            {
                Bold = new Bold()
            };
            fontName = new FontName
            {
                Val = "Calibri"
            };
            fontSize = new FontSize
            {
                Val = 11
            };
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);
            fonts.Count = (uint)fonts.ChildElements.Count;

            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill
            {
                PatternType = PatternValues.None
            };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Gray125
            };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);


            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            var border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            var cellStyleFormats = new CellStyleFormats();
            var cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellStyleFormats.Append(cellFormat);
            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;


            var cellFormats = new CellFormats();

            cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };
            cellFormats.Append(cellFormat);

            var numberingFormats = new NumberingFormats();

            numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            stylesheet.Append(numberingFormats);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);

            var cellStyles = new CellStyles(
                new CellStyle()
                {
                    Name = "Normal",
                    FormatId = 0,
                    BuiltinId = 0,
                }
                );

            cellStyles.Count = (uint)cellStyles.ChildElements.Count;
            stylesheet.Append(cellStyles);

            var differentialFormats = new DifferentialFormats
            {
                Count = 0
            };
            stylesheet.Append(differentialFormats);

            var tableStyles = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            stylesheet.Append(tableStyles);
            return stylesheet;
        }

        private static void SaveCustomStylesheet(WorkbookPart workbookPart)
        {
            //get a copy of the default excel style sheet then add additional styles to it
            var stylesheet = CreateDefaultStylesheet();

            // ***************************** Fills *********************************
            var fills = stylesheet.Fills;

            //header fills background color
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("C8EEFF") },
                    //BackgroundColor = new BackgroundColor() { Indexed = 64 }
                }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            // *************************** numbering formats ***********************
            var numberingFormats = stylesheet.NumberingFormats;
            //number less than 164 is reserved by excel for default formats
            uint iExcelIndex = 165;
            var dateTimeFormat = new NumberingFormat
            {
                NumberFormatId = iExcelIndex++,
                FormatCode = @"[$-409]m/d/yyyy\ h:mm\ AM/PM;@"
            };
            numberingFormats.Append(dateTimeFormat);

            var decimalFormat = new NumberingFormat
            {
                NumberFormatId = iExcelIndex++,
                FormatCode = @"#,##0.00"
            };
            numberingFormats.Append(decimalFormat);

            numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;

            //************************** begin cell formats ***********************************
            var cellFormats = stylesheet.CellFormats; //this should already contain a default StyleIndex of 0

            // Date time format is defined as StyleIndex = 1
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = dateTimeFormat.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            var currentCellFormatIndex = 1;

            // Number format is defined as StyleIndex = 2
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = decimalFormat.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            currentCellFormatIndex++;

            // Header format is defined as StyleIndex = 3
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 2,
                ApplyFill = true,
                BorderId = 0,
                FormatId = 0
            });
            currentCellFormatIndex++;

            // Rotated and centered header is defined as StyleIndex = 4
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 2,
                ApplyFill = true,
                BorderId = 0,
                FormatId = 0,
                Alignment = new Alignment
                {
                    TextRotation = (UInt32Value)90,
                    Horizontal = HorizontalAlignmentValues.Center
                },
            });
            currentCellFormatIndex++;

            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            //************************** end cell formats ***********************************

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = stylesheet;
            style.Save();
        }
    }
}

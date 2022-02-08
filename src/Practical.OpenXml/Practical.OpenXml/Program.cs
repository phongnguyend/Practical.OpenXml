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

            var headerList = new string[] { "Header 1", "Header 2", "Header 3", "Header 4" };

            var stopWatch = new Stopwatch();

            using (var spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                var workbookPart = spreadSheet.AddWorkbookPart();

                SaveCustomStylesheet(workbookPart);


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
                        a[i, j] = new string(
                            Enumerable.Repeat(chars, 5)
                                        .Select(s => s[random.Next(s.Length)])
                                        .ToArray());
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
                        //header formatting attribute.  This will create a <c> element with s=2 as its attribute
                        //s stands for styleindex
                        var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "2") }.ToList();
                        writer.WriteSharedStringCellValue(headerList[i], sharedStringData, attributes);

                    }
                    writer.WriteEndElement(); //end of Row tag

                    for (int i = 0; i < 100; i++)
                    {
                        writer.WriteStartElement(new Row());
                        for (int j = 0; j < 4; j++)
                        {
                            writer.WriteInlineStringCellValue(a[i, j]);
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
            var ss = new Stylesheet();

            var fts = new Fonts();
            var ft = new Font();
            var ftn = new FontName
            {
                Val = "Calibri"
            };
            var ftsz = new FontSize
            {
                Val = 11
            };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            ft = new Font
            {
                Bold = new Bold()
            };
            ftn = new FontName
            {
                Val = "Calibri"
            };
            ftsz = new FontSize
            {
                Val = 11
            };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

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

            var csfs = new CellStyleFormats();
            var cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;


            var cfs = new CellFormats();

            cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };
            cfs.Append(cf);

            var nfs = new NumberingFormats();

            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            var css = new CellStyles(
                new CellStyle()
                {
                    Name = "Normal",
                    FormatId = 0,
                    BuiltinId = 0,
                }
                );

            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            var dfs = new DifferentialFormats
            {
                Count = 0
            };
            ss.Append(dfs);

            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            ss.Append(tss);
            return ss;
        }

        private static void SaveCustomStylesheet(WorkbookPart workbookPart)
        {

            //get a copy of the default excel style sheet then add additional styles to it
            var stylesheet = CreateDefaultStylesheet();

            // ***************************** Fills *********************************
            var fills = stylesheet.Fills;

            //header fills background color
            var fill = new Fill();
            var patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("C8EEFF") };
            //patternFill.BackgroundColor = new BackgroundColor() { Indexed = 64 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            // *************************** numbering formats ***********************
            var nfs = stylesheet.NumberingFormats;
            //number less than 164 is reserved by excel for default formats
            uint iExcelIndex = 165;
            var nf = new NumberingFormat
            {
                NumberFormatId = iExcelIndex++,
                FormatCode = @"[$-409]m/d/yy\ h:mm\ AM/PM;@"
            };
            nfs.Append(nf);

            nfs.Count = (uint)nfs.ChildElements.Count;

            //************************** cell formats ***********************************
            var cfs = stylesheet.CellFormats; //this should already contain a default StyleIndex of 0

            var cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            // Date time format is defined as StyleIndex = 1
            cfs.Append(cf);

            cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 2,
                ApplyFill = true,
                BorderId = 0,
                FormatId = 0
            };
            // Header format is defined as StyleINdex = 2
            cfs.Append(cf);


            cfs.Count = (uint)cfs.ChildElements.Count;

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = stylesheet;
            style.Save();
        }
    }
}

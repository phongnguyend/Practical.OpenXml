using ExcelDataReader;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Practical.ExcelDataReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            Console.WriteLine($"Starting: {GetTotalMegaBytesOfMemoryUsed()} MB");

            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = Path.Combine(path, @"Practical.ExcelDataReader.xlsx");
            using (var fileStream = File.OpenRead(fileName))
            {
                Console.WriteLine($"Loading: {GetTotalMegaBytesOfMemoryUsed()} MB");

                Stopwatch stopwatch = Stopwatch.StartNew();
                var rows = CountRumberOfRecords(fileStream);
                ReadDataByAddresses(fileStream, new[] { "G4", "I10" });
                stopwatch.Stop();
                Console.WriteLine($"Rows Count: {rows}, Took: {stopwatch.Elapsed.TotalSeconds} seconds.");
                Console.WriteLine($"Loaded: {GetTotalMegaBytesOfMemoryUsed()} MB");
            }

            Console.WriteLine($"Ending: {GetTotalMegaBytesOfMemoryUsed()} MB");

            Console.ReadLine();
        }

        private static int CountRumberOfRecords(Stream stream, IEnumerable<string>? ignoredTabs = null)
        {
            int numberOfRecords = 0;

            var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                if (reader.VisibleState == "hidden")
                {
                    continue;
                }

                var sheetName = reader.Name;
                if (ignoredTabs != null && sheetName.In(ignoredTabs))
                {
                    continue;
                }

                numberOfRecords += reader.RowCount;

            } while (reader.NextResult());

            return numberOfRecords;
        }

        private static IEnumerable<ExcelCell> ReadDataByAddresses(Stream stream, IEnumerable<string> addresses, IEnumerable<string>? ignoredTabs = null)
        {
            var cells = addresses.Select(x => new ExcelCell { Address = x }).ToList();

            foreach (var cell in cells)
            {
                var matched = Regex.Match(cell.Address, @"^(\D+)(\d+)$");
                var columnName = matched.Groups[1].Value;
                cell.ColumnIndex = ExcelConverter.ColumNameToIndex(columnName);
                cell.RowIndex = int.Parse(matched.Groups[2].Value) - 1;
            }

            var maxRowIndex = cells.Max(x => x.RowIndex);

            var reader = ExcelReaderFactory.CreateReader(stream);

            var result = new List<ExcelCell>();

            do
            {
                if (reader.VisibleState == "hidden")
                {
                    continue;
                }

                var sheetName = reader.Name;
                if (ignoredTabs != null && sheetName.In(ignoredTabs))
                {
                    continue;
                }

                var rowIndex = 0;
                while (reader.Read())
                {
                    foreach (var cell in cells.Where(x => x.RowIndex == rowIndex))
                    {
                        result.Add(new ExcelCell
                        {
                            Address = cell.Address,
                            ColumnIndex = cell.ColumnIndex,
                            RowIndex = cell.RowIndex,
                            SheetName = reader.Name,
                            Value = reader.GetValue(cell.ColumnIndex)?.ToString(),
                        });
                    }

                    if (rowIndex >= maxRowIndex)
                    {
                        break;
                    }

                    rowIndex++;
                }

            } while (reader.NextResult());

            return result;
        }

        private static string GetTotalMegaBytesOfMemoryUsed()
        {
            long totalBytesOfMemoryUsed = Environment.WorkingSet / 1024 / 1024;
            return totalBytesOfMemoryUsed.ToString("N0");
        }
    }
}
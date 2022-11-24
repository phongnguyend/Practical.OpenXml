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
                FindHeaderRow(fileStream, new[] { "Insured Name", "Invoice Date" });

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
                if (reader.VisibleState != "visible") // visible, hidden, veryhidden
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
                cell.ColumnIndex = ExcelConverter.ColumnNameToIndex(columnName);
                cell.RowIndex = int.Parse(matched.Groups[2].Value) - 1;
            }

            var maxRowIndex = cells.Max(x => x.RowIndex);

            var reader = ExcelReaderFactory.CreateReader(stream);

            var result = new List<ExcelCell>();

            do
            {
                if (reader.VisibleState != "visible") // visible, hidden, veryhidden
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

        private static IEnumerable<ExcelCell> FindHeaderRow(IExcelDataReader reader, IEnumerable<string> headerNames, ref int rowIndex, int? fixedHeaderRowNumber = null)
        {
            var mergeCells = reader.MergeCells;
            var mergedHeaders = new List<ExcelCell>();

            while (reader.Read())
            {
                var headers = new List<ExcelCell>();
                for (int columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                {
                    var headerValue = reader.GetValue(columnIndex)?.ToString();
                    var header = new ExcelCell { Value = headerValue, RowIndex = rowIndex, ColumnIndex = columnIndex, SheetName = reader.Name, Address = $"{ExcelConverter.ColumnIndexToName(columnIndex)}{rowIndex + 1}" };
                    headers.Add(header);
                    var tempRowIndex = rowIndex;
                    if (mergeCells != null && mergeCells.Any(c => c.FromRow == tempRowIndex && c.FromColumn == columnIndex))
                    {
                        mergedHeaders.Add(header);
                    }
                }

                // Ignore if the headers are vertical merged cells
                if (mergeCells != null && mergeCells.Any(c => headers.Any(h => h.RowIndex == c.FromRow && h.ColumnIndex == c.FromColumn && c.FromColumn == c.ToColumn)))
                {
                    rowIndex++;
                    continue;
                }

                headers.ForEach(h =>
                {
                    var mergeCell = mergeCells != null ? mergeCells.FirstOrDefault(c => h.RowIndex >= c.FromRow && h.RowIndex <= c.ToRow && h.ColumnIndex >= c.FromColumn && h.ColumnIndex <= c.ToRow) : null;
                    if (mergeCell != null)
                    {
                        h.Value = mergedHeaders.First(x => x.RowIndex == mergeCell.FromRow && x.ColumnIndex == mergeCell.FromColumn).Value;
                    }
                });

                var finalCandidates = headers.GroupBy(h => h.Value, StringComparer.InvariantCultureIgnoreCase)
                           .SelectMany(g => g.Select((h, k) => new ExcelCell
                           {
                               Value = (h.Value ?? string.Empty).Trim(),
                               RowIndex = h.RowIndex,
                               ColumnIndex = h.ColumnIndex,
                               SheetName = h.SheetName,
                               Address = h.Address,
                           }));

                if (fixedHeaderRowNumber.HasValue && rowIndex < fixedHeaderRowNumber.Value - 1)
                {
                    rowIndex++;
                    continue;
                }

                if (headerNames.All(x => finalCandidates.Any(y => y.Value == x)))
                {
                    return finalCandidates.Where(x => !string.IsNullOrWhiteSpace(x.Value));
                }

                if (rowIndex >= 99 || (fixedHeaderRowNumber.HasValue && rowIndex == fixedHeaderRowNumber.Value - 1))
                {
                    return new List<ExcelCell>();
                }

                rowIndex++;
            }

            return new List<ExcelCell>();
        }

        private static void FindHeaderRow(Stream stream, IEnumerable<string> headerNames, IEnumerable<string>? ignoredTabs = null)
        {
            var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                if (reader.VisibleState != "visible") // visible, hidden, veryhidden
                {
                    continue;
                }

                var sheetName = reader.Name;
                if (ignoredTabs != null && sheetName.In(ignoredTabs))
                {
                    continue;
                }

                int rowIndex = 0;
                var headers = FindHeaderRow(reader, headerNames, ref rowIndex);

            } while (reader.NextResult());
        }
    }
}
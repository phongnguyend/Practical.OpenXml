using ExcelDataReader;
using System.Diagnostics;

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

        private static string GetTotalMegaBytesOfMemoryUsed()
        {
            long totalBytesOfMemoryUsed = Environment.WorkingSet / 1024 / 1024;
            return totalBytesOfMemoryUsed.ToString("N0");
        }
    }
}
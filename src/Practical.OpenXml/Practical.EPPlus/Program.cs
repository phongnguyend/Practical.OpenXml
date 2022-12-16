namespace Practical.EPPlus
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ConfigurationEntryExcelReader reader = new ConfigurationEntryExcelReader();

            using (var fileStream = File.OpenRead("ConfigurationEntries.xlsx"))
            {
                var entries = reader.Read(fileStream);
            }

            ConfigurationEntryExcelWriter writer = new ConfigurationEntryExcelWriter();

            using (var fileStream = File.OpenWrite("ConfigurationEntries1.xlsx"))
            {
                writer.Write(new List<ConfigurationEntry> {
                    new ConfigurationEntry
                    {
                        Key = "Key1",
                        Value = "Value 1",
                    },
                    new ConfigurationEntry
                    {
                        Key = "Key2",
                        Value = "Value 2",
                    },
                    new ConfigurationEntry
                    {
                        Key = "Key5",
                        Value = "Value 5",
                    },
                }, fileStream);
            }

            using (var fileStream = File.OpenRead("ConfigurationEntries1.xlsx"))
            {
                var entries = reader.Read(fileStream);
            }
        }
    }
}
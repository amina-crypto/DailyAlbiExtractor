using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using System.Threading.Tasks;

namespace DailyAlbiExtractor
{
    public class Program
    {
        public static void Main(string[] args)
        {
            MainAsync().GetAwaiter().GetResult();
        }

        private static async Task MainAsync()
        {
            Console.WriteLine($"Starting execution at {DateTime.Now}");
            // Ensure data folder exists
            Directory.CreateDirectory(DataFetcher.DataFolder);
            Console.WriteLine($"Data folder created/verified: {DataFetcher.DataFolder}");

            // Fetch current data
            var fetcher = new DataFetcher();
            var currentData = fetcher.FetchAllData();
            Console.WriteLine($"Fetched {currentData.Count} items");

            // Generate filenames with today's date
            var today = DateTime.Now.ToString("yyyyMMdd");
            var fullExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}.xlsx");
           
            // Save full data to Excel
            var excelHandler = new ExcelHandler();
            excelHandler.SaveToExcel(currentData, fullExcelPath);

          
            excelHandler.DownloadExcelFile(fullExcelPath);
           
            Console.WriteLine($"Execution completed at {DateTime.Now}");
        }
    }
}
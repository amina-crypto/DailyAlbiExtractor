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


            var fetcher = new DataFetcher();

            // Fetch local data
            var localData = fetcher.FetchAllDataFromExcel();
            Console.WriteLine($"Fetched {localData.Count} items");

            // Fetch current data
            var currentData = fetcher.FetchAllDataFromApi();
            Console.WriteLine($"Fetched {currentData.Count} items");

            // Generate filenames with today's date
            var today = DateTime.Now.ToString("yyyyMMdd");
            var fullExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}.xlsx");
            var changesExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}_Changes.xlsx");

            // Save full data and changes to Excel
            var excelHandler = new ExcelHandler();
            excelHandler.SaveToExcel(currentData, fullExcelPath);

            // Download both Excel files
            excelHandler.DownloadExcelFile(fullExcelPath);
            if (File.Exists(changesExcelPath))
            {
                excelHandler.DownloadExcelFile(changesExcelPath);
            }
            else
            {
                Console.WriteLine($"Changes file not found at: {changesExcelPath}");
            }

            Console.WriteLine($"Execution completed at {DateTime.Now}");
        }
    }
}
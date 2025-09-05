using System;
using System.Collections.Generic;
using System.IO;
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
            Console.WriteLine($"Starting execution at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

            // Ensure data folder exists
            Directory.CreateDirectory(DataFetcher.DataFolder);
            Console.WriteLine($"Data folder: {DataFetcher.DataFolder}");

            var fetcher = new DataFetcher();

            // Load previous (optional, just logs)
            var localData = fetcher.FetchAllDataFromExcel();
            Console.WriteLine($"Local (prev) items: {localData.Count}");

            // Fetch current data from API
            var currentData = fetcher.FetchAllDataFromApi();
            Console.WriteLine($"API items: {currentData.Count}");

            // File names for today
            var today = DateTime.Now.ToString("yyyyMMdd");
            var fullExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}.xlsx");
            var changesExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}_Changes.xlsx");

            // Save Excel files (ExcelHandler also makes the Changes workbook)
            var excelHandler = new ExcelHandler();
            excelHandler.SaveToExcel(currentData, fullExcelPath);

            // Optional: copy to Downloads
            excelHandler.DownloadExcelFile(fullExcelPath);
            if (File.Exists(changesExcelPath))
                excelHandler.DownloadExcelFile(changesExcelPath);
            else
                Console.WriteLine($"[INFO] Changes file not found at: {changesExcelPath} (first run or no diffs)");

            // Ask to send via Outlook desktop (COM)
            Console.Write("Send email now with the Excel files via Outlook desktop? (y/n): ");
            var answer = Console.ReadLine()?.Trim().ToLowerInvariant();
            if (answer == "y" || answer == "yes")
            {
                Console.Write("Recipient emails (comma-separated): ");
                var recipientsRaw = Console.ReadLine();
                var recipients = (recipientsRaw ?? string.Empty).Split(',');

                // Build attachment list (only files that exist)
                var attachments = new List<string>();
                if (File.Exists(fullExcelPath)) { Console.WriteLine($"Attaching: {fullExcelPath}"); attachments.Add(fullExcelPath); }
                else Console.WriteLine($"[WARN] Missing: {fullExcelPath}");
                if (File.Exists(changesExcelPath)) { Console.WriteLine($"Attaching: {changesExcelPath}"); attachments.Add(changesExcelPath); }
                else Console.WriteLine($"[WARN] Missing: {changesExcelPath}");

                if (attachments.Count == 0)
                {
                    Console.WriteLine("[INFO] No Excel files found to attach. Email will not be sent.");
                }
                else
                {
                    try
                    {
                        var subject = $"Daily API Extract - {today}";
                        var body =
@"Attached are:
• Full extract (FullData_yyyyMMdd.xlsx)
• Changes vs previous (FullData_yyyyMMdd_Changes.xlsx)
Regards.";

                        // Uses your Outlook profile (Microsoft.Office.Interop.Outlook)
                        OutlookComSender.Send(subject, body, attachments, recipients);
                        Console.WriteLine("✅ Email sent via Outlook desktop.");
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine("❌ ERROR sending via Outlook COM:");
                        Console.WriteLine(ex.ToString());
                        if (ex.InnerException != null) Console.WriteLine("Inner: " + ex.InnerException.Message);
                    }
                }
            }

            Console.WriteLine($"Execution completed at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        }
    }
}

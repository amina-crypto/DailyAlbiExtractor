using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Configuration;

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
            // Fetch current data
            var fetcher = new DataFetcher();
            var currentData = fetcher.FetchAllData();

            // Generate filenames with today's date
            var today = DateTime.Now.ToString("yyyyMMdd");
            var fullExcelPath = System.IO.Path.Combine(DataFetcher.DataFolder, $"FullData_{today}.xlsx");
            var changesExcelPath = System.IO.Path.Combine(DataFetcher.DataFolder, $"Changes_{today}.xlsx");

            // Save full data to Excel
            var excelHandler = new ExcelHandler();
            excelHandler.SaveToExcel(currentData, fullExcelPath);

            // Find previous data file (latest before today)
            var previousFile = excelHandler.GetLatestPreviousFile();
            List<ApiItem> previousData = null;
            if (previousFile != null)
            {
                previousData = excelHandler.LoadFromExcel(previousFile);
            }

            // Detect changes/additions
            var detector = new ChangeDetector();
            var changes = detector.DetectChanges(previousData ?? new List<ApiItem>(), currentData);

            // Save changes to Excel if any
            if (changes.Any())
            {
                excelHandler.SaveToExcel(changes, changesExcelPath);
            }

            //// Send email with attachments
            //var emailSender = new EmailSender(
            //    ConfigurationManager.AppSettings["SmtpServer"],
            //    int.Parse(ConfigurationManager.AppSettings["SmtpPort"]),
            //    ConfigurationManager.AppSettings["SmtpUsername"],
            //    ConfigurationManager.AppSettings["SmtpPassword"],
            //    ConfigurationManager.AppSettings["FromEmail"],
            //    ConfigurationManager.AppSettings["ToEmails"].Split(',')
            //);
            //emailSender.SendEmail(changes.Any() ? new[] { fullExcelPath, changesExcelPath } : new[] { fullExcelPath });
        }
    }
}
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

            // Fetch local data (optional, used by your code for comparison logs)
            var localData = fetcher.FetchAllDataFromExcel();
            Console.WriteLine($"Local (prev) items: {localData.Count}");

            // Fetch current data from API
            var currentData = fetcher.FetchAllDataFromApi();
            Console.WriteLine($"API items: {currentData.Count}");

            // Generate filenames with today's date
            var today = DateTime.Now.ToString("yyyyMMdd");
            var fullExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}.xlsx");
            var changesExcelPath = Path.Combine(DataFetcher.DataFolder, $"FullData_{today}_Changes.xlsx");

            // Save full data and changes to Excel (your ExcelHandler already creates both files)
            var excelHandler = new ExcelHandler();
            excelHandler.SaveToExcel(currentData, fullExcelPath);

            // Optionally copy to Downloads as you already do
            excelHandler.DownloadExcelFile(fullExcelPath);
            if (File.Exists(changesExcelPath))
                excelHandler.DownloadExcelFile(changesExcelPath);
            else
                Console.WriteLine($"Changes file not found at: {changesExcelPath}");

            // Ask if user wants to email the files now
            Console.Write("Send email now with the Excel files? (y/n): ");
            var answer = Console.ReadLine()?.Trim().ToLowerInvariant();
            if (answer == "y" || answer == "yes")
            {
                // Collect SMTP details & recipients through console
                Console.Write("SMTP server (e.g., smtp.gmail.com): ");
                var smtpServer = Console.ReadLine();

                Console.Write("SMTP port (e.g., 587): ");
                var smtpPortStr = Console.ReadLine();
                int smtpPort = 587;
                int.TryParse(smtpPortStr, out smtpPort);

                Console.Write("SMTP username (login email/username): ");
                var smtpUser = Console.ReadLine();

                Console.Write("SMTP password (input hidden): ");
                var smtpPass = ReadPassword();

                Console.Write("From email (the sender): ");
                var fromEmail = Console.ReadLine();

                Console.Write("Recipient emails (comma-separated): ");
                var recipientsRaw = Console.ReadLine();
                var recipients = (recipientsRaw ?? string.Empty).Split(',');

                // Build attachments list (include only files that exist)
                var attachments = new List<string>();
                if (File.Exists(fullExcelPath)) attachments.Add(fullExcelPath);
                if (File.Exists(changesExcelPath)) attachments.Add(changesExcelPath);

                if (attachments.Count == 0)
                {
                    Console.WriteLine("No Excel files found to attach. Email will not be sent.");
                }
                else
                {
                    try
                    {
                        var sender = new EmailSender(
                            smtpServer: smtpServer,
                            smtpPort: smtpPort,
                            smtpUsername: smtpUser,
                            smtpPassword: smtpPass,
                            fromEmail: fromEmail,
                            toEmails: recipients
                        );

                        var subject = $"Daily API Extract - {today}";
                        var body =
@"Attached are:
• Full extract (FullData_yyyyMMdd.xlsx)
• Changes vs previous (FullData_yyyyMMdd_Changes.xlsx)
Regards.";

                        sender.SendEmail(subject, body, attachments);
                        Console.WriteLine("Email sent successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to send email: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"Execution completed at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        }

        /// <summary>
        /// Read a password from console without echoing characters.
        /// </summary>
        private static string ReadPassword()
        {
            var sb = new System.Text.StringBuilder();
            ConsoleKeyInfo keyInfo;
            while ((keyInfo = Console.ReadKey(true)).Key != ConsoleKey.Enter)
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (sb.Length > 0)
                    {
                        sb.Length--;
                        // erase asterisk from console
                        Console.Write("\b \b");
                    }
                }
                else if (!char.IsControl(keyInfo.KeyChar))
                {
                    sb.Append(keyInfo.KeyChar);
                    Console.Write("*");
                }
            }
            Console.WriteLine();
            return sb.ToString();
        }
    }
}

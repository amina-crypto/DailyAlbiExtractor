using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;

namespace DailyAlbiExtractor
{
    public class DataFetcher
    {
        public const string DataFolder = @"C:\DataExtracts"; // Changed to absolute path for consistency
        private const string ApiBaseUrl = "https://politicheattive.lavoro.gov.it/albi-informatici_service/public/UI/search/paged";
        private const int SectionId = 1;
        private const string OrderBy = "id";
        private const bool Ascendente = true;
        private const int Limit = 1000; // Covers all expected records
        private readonly ExcelHandler _excelHandler;

        public DataFetcher()
        {
            _excelHandler = new ExcelHandler();
        }

        public List<ApiItem> FetchAllDataFromExcel()
        {
            var allItems = new List<ApiItem>();
            var today = DateTime.Now.ToString("yyyyMMdd");
            var excelFilePath = Path.Combine(DataFolder, $"FullData_{today}_prev.xlsx");

            try
            {
                Console.WriteLine($"Reading data from Excel file: {excelFilePath}");
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine($"Excel file not found: {excelFilePath}");
                    return allItems; // Return empty list if file doesn't exist
                }

                // Use ExcelHandler to read data from the Excel file
                allItems = _excelHandler.CaricaDaExcel(excelFilePath);
                Console.WriteLine($"Total items read from Excel: {allItems.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message} - StackTrace: {ex.StackTrace}");
            }

            Console.WriteLine($"Total items fetched from Excel: {allItems.Count}");
            return allItems;
        }

        public List<ApiItem> FetchAllDataFromApi()
        {
            var allItems = new List<ApiItem>();
            int offset = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/json");
                while (true)
                {
                    try
                    {
                        string url = $"{ApiBaseUrl}?idSezione={SectionId}&orderBy={OrderBy}&ascendente={Ascendente}&offset={offset}&limit={Limit}";
                        Console.WriteLine($"Fetching data from: {url}");
                        string jsonString = client.DownloadString(url);
                        Console.WriteLine($"Received JSON length: {jsonString.Length}");
                        var apiResponse = JsonConvert.DeserializeObject<ApiResponse>(jsonString);
                        if (apiResponse == null || apiResponse.Content == null)
                        {
                            Console.WriteLine("API response or content is null, exiting loop.");
                            break;
                        }
                        allItems.AddRange(apiResponse.Content);
                        if (apiResponse.Last || apiResponse.NumberOfElements == 0 || apiResponse.NumberOfElements < Limit)
                        {
                            Console.WriteLine($"Last page reached. Total items: {allItems.Count}");
                            break;
                        }
                        offset += Limit;
                    }
                    catch (WebException ex)
                    {
                        Console.WriteLine($"Error fetching data: {ex.Message} - Status: {ex.Status}");
                        break;
                    }
                    catch (JsonException ex)
                    {
                        Console.WriteLine($"Error deserializing JSON: {ex.Message}");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unexpected error: {ex.Message} - StackTrace: {ex.StackTrace}");
                        break;
                    }
                }
            }
            Console.WriteLine($"Total items fetched from API: {allItems.Count}");
            return allItems;
        }
    }
}
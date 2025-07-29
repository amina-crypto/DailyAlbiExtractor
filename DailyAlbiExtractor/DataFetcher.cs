using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;

namespace DailyAlbiExtractor
{
    public class DataFetcher
    {
        public const string DataFolder = "DataExtracts";
        private const string ApiBaseUrl = "https://politicheattive.lavoro.gov.it/albi-informatici_service/public/UI/search/paged";
        private const int SectionId = 1;
        private const string OrderBy = "id";
        private const bool Ascendente = true;
        private const int Limit = 1000; // Set to 1000 as per your URL, which should cover all 218 records

        /// <summary>
        /// Fetches all data from the API, handling pagination by incrementing the offset.
        /// </summary>
        /// <returns>A list of ApiItem objects containing all fetched data.</returns>
        public List<ApiItem> FetchAllData()
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
                            break; // Exit if response or content is null
                        }

                        allItems.AddRange(apiResponse.Content);

                        // Stop if we've reached the last page or no more data
                        if (apiResponse.Last || apiResponse.NumberOfElements == 0 || apiResponse.NumberOfElements < Limit)
                        {
                            Console.WriteLine($"Last page reached. Total items: {allItems.Count}");
                            break;
                        }

                        offset += Limit; // Move to the next page
                    }
                    catch (WebException ex)
                    {
                        Console.WriteLine($"Error fetching data: {ex.Message} - Status: {ex.Status}");
                        break;
                    }
                    catch (JsonException ex)
                    {
                        Console.WriteLine($"Error deserializing JSON: {ex.Message}"); // Removed Path
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unexpected error: {ex.Message} - StackTrace: {ex.StackTrace}");
                        break;
                    }
                }
            }

            Console.WriteLine($"Total items fetched: {allItems.Count}");
            return allItems;
        }
    }
}
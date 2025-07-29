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
        private const int PageSize = 50; // Matches the API's pageSize from the response

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
                        string url = $"{ApiBaseUrl}?idSezione={SectionId}&orderBy={OrderBy}&ascendente={Ascendente}&offset={offset}&size={PageSize}";
                        string jsonString = client.DownloadString(url);
                        var apiResponse = JsonConvert.DeserializeObject<ApiResponse>(jsonString);

                        if (apiResponse == null || apiResponse.Content == null)
                        {
                            break; // Exit if response or content is null
                        }

                        allItems.AddRange(apiResponse.Content);

                        // Stop if we've reached the last page
                        if (apiResponse.Last || apiResponse.NumberOfElements < PageSize)
                        {
                            break;
                        }

                        offset += PageSize; // Move to the next page
                    }
                    catch (WebException ex)
                    {
                        Console.WriteLine($"Error fetching data: {ex.Message}");
                        break;
                    }
                    catch (JsonException ex)
                    {
                        Console.WriteLine($"Error deserializing JSON: {ex.Message}");
                        break;
                    }
                }
            }

            return allItems;
        }
    }
}
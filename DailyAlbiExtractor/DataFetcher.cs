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
        private const string ApiUrl = "https://politicheattive.lavoro.gov.it/albi-informatici_service/public/UI/search/paged?idSezione=1&orderBy=id&ascendente=true&offset=0";

        public List<ApiItem> FetchAllData()
        {
            var allItems = new List<ApiItem>();
            using (var client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/json");
                var jsonString = client.DownloadString(ApiUrl);
                var apiResponse = JsonConvert.DeserializeObject<ApiResponse>(jsonString);

                if (apiResponse != null && apiResponse.Content != null)
                {
                    allItems.AddRange(apiResponse.Content);
                }
            }
            return allItems;
        }
    }
}
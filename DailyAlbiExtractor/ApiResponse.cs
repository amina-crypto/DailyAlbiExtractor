using System.Collections.Generic;
using Newtonsoft.Json;

namespace DailyAlbiExtractor
{
    public class ApiResponse
    {
        [JsonProperty("content")]
        public List<ApiItem> Content { get; set; }

        [JsonProperty("numberOfElements")]
        public int NumberOfElements { get; set; }
    }

}

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

            [JsonProperty("first")]
            public bool First { get; set; }

            [JsonProperty("last")]
            public bool Last { get; set; }

            [JsonProperty("number")]
            public int Number { get; set; }

            [JsonProperty("pageable")]
            public Pageable Pageable { get; set; }

            [JsonProperty("size")]
            public int Size { get; set; }

            [JsonProperty("sort")]
            public Sort Sort { get; set; }

            [JsonProperty("totalElements")]
            public int TotalElements { get; set; }

            [JsonProperty("totalPages")]
            public int TotalPages { get; set; }

            [JsonProperty("empty")]
            public bool Empty { get; set; }
        }


}

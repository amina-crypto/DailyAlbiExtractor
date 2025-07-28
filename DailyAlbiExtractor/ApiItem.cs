using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using Newtonsoft.Json;

namespace DailyAlbiExtractor
{
    public class ApiItem
    {
        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("azioneView")]
        public bool AzioneView { get; set; }

        [JsonProperty("codiceFiscale")]
        public string CodiceFiscale { get; set; }

        [JsonProperty("partitaIVA")]
        public string PartitaIVA { get; set; }

        [JsonProperty("ragioneSociale")]
        public string RagioneSociale { get; set; }

        [JsonProperty("statoIscrizione")]
        public string StatoIscrizione { get; set; }

        [JsonProperty("idSezioneAlbo")]
        public int IdSezioneAlbo { get; set; }

        [JsonProperty("descrizioneSezione")]
        public string DescrizioneSezione { get; set; }

        // Add more properties as needed based on the full API response
    }
}


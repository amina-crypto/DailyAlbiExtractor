using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace DailyAlbiExtractor
{
    public class ApiItem
    {
        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("azioneView")]
        public bool AzioneView { get; set; }

        [JsonProperty("azioneEdit")]
        public bool AzioneEdit { get; set; }

        [JsonProperty("azioneDelete")]
        public bool AzioneDelete { get; set; }

        [JsonProperty("azioneStrutturaOrganizzativa")]
        public bool AzioneStrutturaOrganizzativa { get; set; }

        [JsonProperty("azioneGestioneLavoratori")]
        public bool AzioneGestioneLavoratori { get; set; }

        [JsonProperty("azioneGestioneDelegheAmministrative")]
        public bool AzioneGestioneDelegheAmministrative { get; set; }

        [JsonProperty("azioneGestionePoliticheAttive")]
        public bool AzioneGestionePoliticheAttive { get; set; }

        [JsonProperty("idTipologiaPersonaGiuridica")]
        public string IdTipologiaPersonaGiuridica { get; set; }

        [JsonProperty("idStatoSedePatronato")]
        public string IdStatoSedePatronato { get; set; }

        [JsonProperty("idSedePatronato")]
        public string IdSedePatronato { get; set; }

        [JsonProperty("codiceFiscale")]
        public string CodiceFiscale { get; set; }

        [JsonProperty("partitaIVA")]
        public string PartitaIVA { get; set; }

        [JsonProperty("ragioneSociale")]
        public string RagioneSociale { get; set; }

        [JsonProperty("idAteco2007")]
        public string IdAteco2007 { get; set; }

        [JsonProperty("idFormaGiuridica")]
        public string IdFormaGiuridica { get; set; }

        [JsonProperty("codiceREA")]
        public string CodiceREA { get; set; }

        [JsonProperty("numeroSoci")]
        public string NumeroSoci { get; set; }

        [JsonProperty("numeroDipendenti")]
        public string NumeroDipendenti { get; set; }

        [JsonProperty("numeroCollaboratori")]
        public string NumeroCollaboratori { get; set; }

        [JsonProperty("numeroIscrittiLibroSoci")]
        public string NumeroIscrittiLibroSoci { get; set; }

        [JsonProperty("capitaleSociale")]
        public string CapitaleSociale { get; set; }

        [JsonProperty("dataCapitaleSociale")]
        public string DataCapitaleSociale { get; set; }

        [JsonProperty("codiceIscrizioneCCIAA")]
        public string CodiceIscrizioneCCIAA { get; set; }

        [JsonProperty("dataIscrizioneCCIAA")]
        public string DataIscrizioneCCIAA { get; set; }

        [JsonProperty("sitoWeb")]
        public string SitoWeb { get; set; }

        [JsonProperty("iban")]
        public string Iban { get; set; }

        [JsonProperty("email")]
        public string Email { get; set; }

        [JsonProperty("emailPEC")]
        public string EmailPEC { get; set; }

        [JsonProperty("dataCostituzione")]
        public string DataCostituzione { get; set; }

        [JsonProperty("dataCessazione")]
        public string DataCessazione { get; set; }

        [JsonProperty("telefono")]
        public string Telefono { get; set; }

        [JsonProperty("fax")]
        public string Fax { get; set; }

        [JsonProperty("idComuneSedeLegale")]
        public int IdComuneSedeLegale { get; set; }

        [JsonProperty("indirizzoSedeLegale")]
        public string IndirizzoSedeLegale { get; set; }

        [JsonProperty("civicoSedeLegale")]
        public string CivicoSedeLegale { get; set; }

        [JsonProperty("capSedeLegale")]
        public string CapSedeLegale { get; set; }

        [JsonProperty("flagPrivacy")]
        public string FlagPrivacy { get; set; }

        [JsonProperty("idCittadinanza")]
        public string IdCittadinanza { get; set; }

        [JsonProperty("dataInizioValidita")]
        public string DataInizioValidita { get; set; }

        [JsonProperty("dataFineValidita")]
        public string DataFineValidita { get; set; }

        [JsonProperty("utente")]
        public string Utente { get; set; }

        [JsonProperty("codiceSezione")]
        public string CodiceSezione { get; set; }

        [JsonProperty("descrizioneSezione")]
        public string DescrizioneSezione { get; set; }

        [JsonProperty("descrizioneComuneSedeLegale")]
        public string DescrizioneComuneSedeLegale { get; set; }

        [JsonProperty("statoIscrizione")]
        public string StatoIscrizione { get; set; }

        [JsonProperty("idSezioneAlbo")]
        public int IdSezioneAlbo { get; set; }

        [JsonProperty("idAlboIntermediario")]
        public int IdAlboIntermediario { get; set; }

        [JsonProperty("idAlboIntermediarioSezione")]
        public int IdAlboIntermediarioSezione { get; set; }

        [JsonProperty("idTipologiaAutorizzazioneIntermediarioSezione")]
        public int IdTipologiaAutorizzazioneIntermediarioSezione { get; set; }

        [JsonProperty("politica")]
        public string Politica { get; set; }

        [JsonProperty("recapiti")]
        public string Recapiti { get; set; }

        [JsonProperty("desDenominazioneSede")]
        public string DesDenominazioneSede { get; set; }

        [JsonProperty("indirizzoSede")]
        public string IndirizzoSede { get; set; }

        [JsonProperty("sede")]
        public string Sede { get; set; }
    }
}


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;

namespace DailyAlbiExtractor
{
    public class ExcelHandler
    {
        public void SaveToExcel(List<ApiItem> items, string filePath)
        {
            List<ApiItem> previousItems = null;
            string previousFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + "_prev.xlsx");
            string changesFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + "_Changes.xlsx");
            // Load previous data if it exists
            if (File.Exists(previousFilePath))
            {
                previousItems = CaricaDaExcel(previousFilePath);
            }
            // Create workbook for Data sheet
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");
                // Add headers
                var properties = typeof(ApiItem).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = properties[i].Name;
                }
                // Add data rows
                for (int row = 0; row < items.Count; row++)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = properties[col].GetValue(items[row], null);
                        // Handle specific types to avoid formatting issues
                        if (value is DateTime dateValue)
                        {
                            worksheet.Cell(row + 2, col + 1).Value = dateValue;
                            worksheet.Cell(row + 2, col + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                        }
                        else
                        {
                            worksheet.Cell(row + 2, col + 1).Value = value != null ? value.ToString() : null;
                        }
                    }
                }
                // Save the Data workbook
                try
                {
                    // Backup the previous file if it exists
                    if (File.Exists(filePath))
                    {
                        File.Copy(filePath, previousFilePath, true);
                    }
                    workbook.SaveAs(filePath);
                    Console.WriteLine($"Excel file saved to: {filePath} with {items.Count} records");
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine($"Permission error saving Excel file to {filePath}: {ex.Message}");
                    throw;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving Excel file to {filePath}: {ex.Message}");
                    throw;
                }
            }
            // Create separate workbook for Changes sheet
            using (var changesWorkbook = new XLWorkbook())
            {
                var changesSheet = changesWorkbook.Worksheets.Add("Changes");
                // Add headers matching ApiItem properties
                var properties = typeof(ApiItem).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    changesSheet.Cell(1, i + 1).Value = properties[i].Name;
                }

                int changeRow = 2;
                if (previousItems != null)
                {
                    // Identify changes
                    var currentIds = items.Select(i => i.Id).ToHashSet();
                    var previousIds = previousItems.Select(i => i.Id).ToHashSet();
                    // New lines
                    var newItems = items.Where(i => !previousIds.Contains(i.Id));
                    // Missing lines
                    var missingItems = previousItems.Where(i => !currentIds.Contains(i.Id));
                    // Modified lines
                    var modifiedItems = from curr in items
                                        join prev in previousItems on curr.Id equals prev.Id
                                        where !AreItemsEqual(curr, prev)
                                        select new { Current = curr, Previous = prev };
                    foreach (var item in newItems)
                    {
                        for (int c = 0; c < properties.Length; c++)
                        {
                            var value = properties[c].GetValue(item, null);
                            // Handle specific types to avoid formatting issues
                            if (value is DateTime dateValue)
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = dateValue;
                                changesSheet.Cell(changeRow, c + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                            }
                            else if (properties[c].Name == "DescrizioneSezione")
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = "New";
                            }
                            else
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = value != null ? value.ToString() : null;
                            }
                        }
                        changeRow++;
                    }
                    foreach (var item in missingItems)
                    {
                        for (int c = 0; c < properties.Length; c++)
                        {
                            var value = properties[c].GetValue(item, null);
                            // Handle specific types to avoid formatting issues
                            if (value is DateTime dateValue)
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = dateValue;
                                changesSheet.Cell(changeRow, c + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                            }
                            else if (properties[c].Name == "DescrizioneSezione")
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = "Missing";
                            }
                            else
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = value != null ? value.ToString() : null;
                            }
                        }
                        changeRow++;
                    }
                    foreach (var pair in modifiedItems)
                    {
                        for (int c = 0; c < properties.Length; c++)
                        {
                            var value = properties[c].GetValue(pair.Current, null);
                            // Handle specific types to avoid formatting issues
                            if (value is DateTime dateValue)
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = dateValue;
                                changesSheet.Cell(changeRow, c + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                            }
                            else if (properties[c].Name == "DescrizioneSezione")
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = GetChangeDetails(pair.Previous, pair.Current);
                            }
                            else
                            {
                                changesSheet.Cell(changeRow, c + 1).Value = value != null ? value.ToString() : null;
                            }
                        }
                        changeRow++;
                    }
                }
                else
                {
                    changesSheet.Cell(2, 1).Value = "No previous data for comparison";
                }
                // Save the Changes workbook
                try
                {
                    changesWorkbook.SaveAs(changesFilePath);
                    Console.WriteLine($"Changes Excel file saved to: {changesFilePath}");
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine($"Permission error saving Changes Excel file to {changesFilePath}: {ex.Message}");
                    throw;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving Changes Excel file to {changesFilePath}: {ex.Message}");
                    throw;
                }
            }
        }
        private bool AreItemsEqual(ApiItem item1, ApiItem item2)
        {
            var properties = typeof(ApiItem).GetProperties();
            foreach (var prop in properties)
            {
                var value1 = prop.GetValue(item1);
                var value2 = prop.GetValue(item2);
                if (value1 != value2 && (value1 == null || !value1.Equals(value2)))
                {
                    return false;
                }
            }
            return true;
        }
        private string GetChangeDetails(ApiItem oldItem, ApiItem newItem)
        {
            var changes = new List<string>();
            var properties = typeof(ApiItem).GetProperties();
            foreach (var prop in properties)
            {
                var oldValue = prop.GetValue(oldItem);
                var newValue = prop.GetValue(newItem);
                if (oldValue != newValue && (oldValue == null || !oldValue.Equals(newValue)))
                {
                    changes.Add($"{prop.Name}: {oldValue ?? "null"} -> {newValue ?? "null"}");
                }
            }
            return string.Join("; ", changes);
        }
        private string Normalize(string input)
        {
            var stringInput = CleanString(input);

            return string.IsNullOrWhiteSpace(stringInput) ? null : stringInput.Trim();
        }
        private static string CleanString(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            var result = input
                .Trim()
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Replace("\t", " ")                        // Tab -> spazio
                .Replace("\u00A0", " ")                    // Non-breaking space
                .Replace("\u2000", " ")                    // En quad
                .Replace("\u2001", " ")                    // Em quad
                .Replace("\u2002", " ")                    // En space
                .Replace("\u2003", " ")                    // Em space
                .Replace("\u2004", " ")                    // Three-per-em space
                .Replace("\u2005", " ")                    // Four-per-em space
                .Replace("\u2006", " ")                    // Six-per-em space
                .Replace("\u2007", " ")                    // Figure space
                .Replace("\u2008", " ")                    // Punctuation space
                .Replace("\u2009", " ")                    // Thin space
                .Replace("\u200A", " ")                    // Hair space
                .Replace("''", "'")                        // Double single quotes -> single quote
                .Replace("\"\"", "\"")                     // Double double-quotes -> single double-quote  
                .Normalize(NormalizationForm.FormC)
                .Trim();
            // Handle escaped quotes at beginning/end of string
            while (result.StartsWith("''"))
                result = result.Substring(1);
            while (result.EndsWith("''") && result.Length > 2)
                result = result.Substring(0, result.Length - 1);
            // Also remove single quotes at start/end (common in Excel CSV issues)
            while (result.StartsWith("'") && !result.Substring(1).StartsWith("'"))
                result = result.Substring(1);
            while (result.EndsWith("'") && result.Length > 1 && !result.Substring(0, result.Length - 1).EndsWith("'"))
                result = result.Substring(0, result.Length - 1);
            return result;
        }
        public List<ApiItem> CaricaDaExcel(string previousFilePath)
        {
            var list = new List<ApiItem>();
            try
            {
                var workbook = new XLWorkbook(previousFilePath);
                var worksheet = workbook.Worksheet(1);
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip header
                {
                    var tipo = new ApiItem
                    {
                        Id = int.Parse(row.Cell(1).GetValue<string>()),
                        AzioneView = bool.Parse(row.Cell(2).GetValue<string>()),
                        AzioneEdit = bool.Parse(row.Cell(3).GetValue<string>()),
                        AzioneDelete = bool.Parse(row.Cell(4).GetValue<string>()),
                        AzioneStrutturaOrganizzativa = bool.Parse(row.Cell(5).GetValue<string>()),
                        AzioneGestioneLavoratori = bool.Parse(row.Cell(6).GetValue<string>()),
                        AzioneGestioneDelegheAmministrative = bool.Parse(row.Cell(7).GetValue<string>()),
                        AzioneGestionePoliticheAttive = bool.Parse(row.Cell(8).GetValue<string>()),
                        IdTipologiaPersonaGiuridica = Normalize(row.Cell(9).GetValue<string>()),
                        IdStatoSedePatronato = Normalize(row.Cell(10).GetValue<string>()),
                        IdSedePatronato = Normalize(row.Cell(11).GetValue<string>()),
                        CodiceFiscale = row.Cell(12).GetValue<string>(),
                        PartitaIVA = row.Cell(13).GetValue<string>(),
                        RagioneSociale = Normalize(row.Cell(14).GetValue<string>()),
                        IdAteco2007 = Normalize(row.Cell(15).GetValue<string>()),
                        IdFormaGiuridica = Normalize(row.Cell(16).GetValue<string>()),
                        CodiceREA = Normalize(row.Cell(17).GetValue<string>()),
                        NumeroSoci = Normalize(row.Cell(18).GetValue<string>()),
                        NumeroDipendenti = Normalize(row.Cell(19).GetValue<string>()),
                        NumeroCollaboratori = Normalize(row.Cell(20).GetValue<string>()),
                        NumeroIscrittiLibroSoci = Normalize(row.Cell(21).GetValue<string>()),
                        CapitaleSociale = Normalize(row.Cell(22).GetValue<string>()),
                        DataCapitaleSociale = Normalize(row.Cell(23).GetValue<string>()),
                        CodiceIscrizioneCCIAA = Normalize(row.Cell(24).GetValue<string>()),
                        DataIscrizioneCCIAA = Normalize(row.Cell(25).GetValue<string>()),
                        SitoWeb = Normalize(row.Cell(26).GetValue<string>()),
                        Iban = Normalize(row.Cell(27).GetValue<string>()),
                        Email = Normalize(row.Cell(28).GetValue<string>()),
                        EmailPEC = Normalize(row.Cell(29).GetValue<string>()),
                        DataCostituzione = Normalize(row.Cell(30).GetValue<string>()),
                        DataCessazione = Normalize(row.Cell(31).GetValue<string>()),
                        Telefono = Normalize(row.Cell(32).GetValue<string>()),
                        Fax = Normalize(row.Cell(33).GetValue<string>()),
                        IdComuneSedeLegale = int.Parse(row.Cell(34).GetValue<string>()),
                        IndirizzoSedeLegale = Normalize(row.Cell(35).GetValue<string>()),
                        CivicoSedeLegale = Normalize(row.Cell(36).GetValue<string>()),
                        CapSedeLegale = Normalize(row.Cell(37).GetValue<string>()),
                        FlagPrivacy = Normalize(row.Cell(38).GetValue<string>()),
                        IdCittadinanza = Normalize(row.Cell(39).GetValue<string>()),
                        DataInizioValidita = Normalize(row.Cell(40).GetValue<string>()),
                        DataFineValidita = Normalize(row.Cell(41).GetValue<string>()),
                        Utente = Normalize(row.Cell(42).GetValue<string>()),
                        CodiceSezione = row.Cell(43).GetValue<string>(),
                        DescrizioneSezione = row.Cell(44).GetValue<string>(),
                        DescrizioneComuneSedeLegale = row.Cell(45).GetValue<string>(),
                        StatoIscrizione = row.Cell(46).GetValue<string>(),
                        IdSezioneAlbo = int.Parse(row.Cell(47).GetValue<string>()),
                        IdAlboIntermediario = int.Parse(row.Cell(48).GetValue<string>()),
                        IdAlboIntermediarioSezione = int.Parse(row.Cell(49).GetValue<string>()),
                        IdTipologiaAutorizzazioneIntermediarioSezione = int.Parse(row.Cell(50).GetValue<string>()),
                        Politica = row.Cell(51).GetValue<string>(),
                        Recapiti = row.Cell(52).GetValue<string>(),
                        DesDenominazioneSede = row.Cell(53).GetValue<string>(),
                        IndirizzoSede = row.Cell(54).GetValue<string>(),
                        Sede = row.Cell(55).GetValue<string>()
                    };
                    list.Add(tipo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore durante la lettura del file Excel: {ex.Message}");
            }
            return list;
        }
        public List<ApiItem> LoadFromExcel(string filePath)
        {
            var items = new List<ApiItem>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                // Get headers to map columns
                var headers = new Dictionary<string, int>();
                var firstRow = worksheet.Row(1);
                for (int col = 1; col <= firstRow.CellCount(); col++)
                {
                    var header = firstRow.Cell(col).Value.ToString() ?? string.Empty;
                    headers[header] = col;
                }
                // Load rows
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var item = new ApiItem();
                    foreach (var prop in typeof(ApiItem).GetProperties())
                    {
                        if (headers.TryGetValue(prop.Name, out int col))
                        {
                            var cell = row.Cell(col);
                            object cellValue;
                            if (cell.TryGetValue(out cellValue) && cellValue != null && !string.IsNullOrEmpty(cellValue.ToString()))
                            {
                                try
                                {
                                    var valueType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                    var value = Convert.ChangeType(cellValue, valueType);
                                    prop.SetValue(item, value);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Conversion error for property {prop.Name}: {ex.Message}");
                                }
                            }
                        }
                    }
                    items.Add(item);
                }
            }
            return items;
        }
        public void DownloadExcelFile(string sourceFilePath)
        {
            if (File.Exists(sourceFilePath))
            {
                string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                string fileName = Path.GetFileName(sourceFilePath);
                string destinationPath = Path.Combine(downloadsPath, fileName);
                if (!Directory.Exists(downloadsPath))
                {
                    Directory.CreateDirectory(downloadsPath);
                }
                File.Copy(sourceFilePath, destinationPath, true);
                Console.WriteLine($"Excel file downloaded to: {destinationPath}");
            }
            else
            {
                throw new FileNotFoundException($"Source file not found: {sourceFilePath}");
            }
        }
    }
}
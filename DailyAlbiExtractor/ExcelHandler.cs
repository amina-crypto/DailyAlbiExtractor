
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace DailyAlbiExtractor
{
    public class ExcelHandler
    {
        public void SaveToExcel(List<ApiItem> currentData, string filePath)
        {
            // Find the most recent previous file if it exists
            string directory = Path.GetDirectoryName(filePath);
            string pattern = "FullData_*.xlsx";
            var existingFiles = Directory.GetFiles(directory, pattern)
                .Where(f => f != filePath) // Exclude the current file path
                .OrderByDescending(f => f) // Sort by filename descending (yyyyMMdd is sortable)
                .ToList();
            string previousFilePath = existingFiles.FirstOrDefault();
            List<ApiItem> previousItems = filePath != null ? CaricaDaExcel(filePath) : null;

            using (var workbook = new XLWorkbook())
            {
                // Create Data sheet
                var dataSheet = workbook.Worksheets.Add("Data");
                var properties = typeof(ApiItem).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    dataSheet.Cell(1, i + 1).Value = properties[i].Name;
                }
                // Add data rows
                for (int row = 0; row < currentData.Count; row++)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = properties[col].GetValue(currentData[row]);
                        if (value is DateTime dateValue)
                        {
                            dataSheet.Cell(row + 2, col + 1).Value = dateValue;
                            dataSheet.Cell(row + 2, col + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                        }
                        else
                        {
                            dataSheet.Cell(row + 2, col + 1).Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                }
                // Create Changes sheet with same headers as Data sheet
                var changesSheet = workbook.Worksheets.Add("Changes");
                for (int i = 0; i < properties.Length; i++)
                {
                    changesSheet.Cell(1, i + 1).Value = properties[i].Name;
                }
                changesSheet.Cell(1, properties.Length + 1).Value = "ChangeType";

                // Compare current and previous items
                var changes = CompareApiItems(currentData, previousItems);
                int changeRow = 2;
                foreach (var change in changes)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = change.Values[col];
                        if (value is DateTime dateValue)
                        {
                            changesSheet.Cell(changeRow, col + 1).Value = dateValue;
                            changesSheet.Cell(changeRow, col + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                        }
                        else
                        {
                            changesSheet.Cell(changeRow, col + 1).Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                    changesSheet.Cell(changeRow, properties.Length + 1).Value = change.ChangeType;
                    changeRow++;
                }
                if (!changes.Any())
                {
                    changesSheet.Cell(2, 1).Value = "No changes or no previous data";
                    changesSheet.Cell(2, properties.Length + 1).Value = "Info";
                }
                // Save the current file
                try
                {
                    workbook.SaveAs(filePath);
                    Console.WriteLine($"Excel file saved to: {filePath} with {currentData.Count} records");
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
        }



        private List<(ApiItem Item, object[] Values, string ChangeType)> CompareApiItems(List<ApiItem> currentData, List<ApiItem> previousItems)
        {
            var changes = new List<(ApiItem, object[], string)>();
            var properties = typeof(ApiItem).GetProperties();
            var processedIds = new HashSet<int>(); // To ensure each ID appears only once

            if (previousItems == null || !previousItems.Any())
            {
                return changes; // Return empty list if no previous data - no changes to report
            }

            // Iterate over current items to find new or modified items
            foreach (var currentItem in currentData)
            {
                bool foundMatch = false;
                foreach (var prevItem in previousItems)
                {
                    if (prevItem.Id == currentItem.Id)
                    {
                        foundMatch = true;
                        if (!processedIds.Contains(currentItem.Id)) // Process only if not already handled
                        {
                            var values = new object[properties.Length];
                            bool hasChanges = false;
                            for (int i = 0; i < properties.Length; i++)
                            {
                                var prop = properties[i];
                                var oldValue = prop.GetValue(prevItem);
                                var newValue = prop.GetValue(currentItem);
                                if (oldValue != newValue && (oldValue == null || !oldValue.Equals(newValue)))
                                {
                                    values[i] = newValue;
                                    hasChanges = true;
                                }
                                else
                                {
                                    values[i] = null; // Unchanged properties are null (will be blank in Excel)
                                }
                            }
                            if (hasChanges)
                            {
                                changes.Add((currentItem, values, "Modified"));
                                processedIds.Add(currentItem.Id); // Mark as processed
                            }
                        }
                        break;
                    }
                }
                if (!foundMatch && !processedIds.Contains(currentItem.Id)) // New item, add only once
                {
                    var values = new object[properties.Length];
                    for (int i = 0; i < properties.Length; i++)
                    {
                        values[i] = properties[i].GetValue(currentItem);
                    }
                    changes.Add((currentItem, values, "New"));
                    processedIds.Add(currentItem.Id); // Mark as processed
                }
            }

            // Iterate over previous items to check for deleted items, including the last non-matching ID
            foreach (var prevItem in previousItems)
            {
                if (!currentData.Any(c => c.Id == prevItem.Id) && !processedIds.Contains(prevItem.Id))
                {
                    var values = new object[properties.Length];
                    for (int i = 0; i < properties.Length; i++)
                    {
                        values[i] = properties[i].GetValue(prevItem);
                    }
                    changes.Add((prevItem, values, "Deleted"));
                    processedIds.Add(prevItem.Id); // Mark as processed
                }
            }

            return changes;
        }




        public List<ApiItem> CaricaDaExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var list = new List<ApiItem>();
            foreach (var row in worksheet.RowsUsed().Skip(1)) // Salta header
            {
                var tipo = new ApiItem();
                var value1 = row.Cell(1).GetValue<string>();
                tipo.Id = !string.IsNullOrEmpty(value1) ? int.Parse(value1) : 0;

                var value2 = row.Cell(2).GetValue<string>();
                tipo.AzioneView = !string.IsNullOrEmpty(value2) ? bool.Parse(value2) : false;

                var value3 = row.Cell(3).GetValue<string>();
                tipo.AzioneEdit = !string.IsNullOrEmpty(value3) ? bool.Parse(value3) : false;

                var value4 = row.Cell(4).GetValue<string>();
                tipo.AzioneDelete = !string.IsNullOrEmpty(value4) ? bool.Parse(value4) : false;

                var value5 = row.Cell(5).GetValue<string>();
                tipo.AzioneStrutturaOrganizzativa = !string.IsNullOrEmpty(value5) ? bool.Parse(value5) : false;

                var value6 = row.Cell(6).GetValue<string>();
                tipo.AzioneGestioneLavoratori = !string.IsNullOrEmpty(value6) ? bool.Parse(value6) : false;

                var value7 = row.Cell(7).GetValue<string>();
                tipo.AzioneGestioneDelegheAmministrative = !string.IsNullOrEmpty(value7) ? bool.Parse(value7) : false;

                var value8 = row.Cell(8).GetValue<string>();
                tipo.AzioneGestionePoliticheAttive = !string.IsNullOrEmpty(value8) ? bool.Parse(value8) : false;

                var value9 = row.Cell(9).GetValue<string>();
                tipo.IdTipologiaPersonaGiuridica = !string.IsNullOrEmpty(value9) ? (int?)int.Parse(value9) : null;

                var value10 = row.Cell(10).GetValue<string>();
                tipo.IdStatoSedePatronato = !string.IsNullOrEmpty(value10) ? (int?)int.Parse(value10) : null;

                var value11 = row.Cell(11).GetValue<string>();
                tipo.IdSedePatronato = !string.IsNullOrEmpty(value11) ? (int?)int.Parse(value11) : null;

                tipo.CodiceFiscale = row.Cell(12).GetValue<string>() ?? string.Empty;

                tipo.PartitaIVA = row.Cell(13).GetValue<string>() ?? string.Empty;

                tipo.RagioneSociale = row.Cell(14).GetValue<string>() ?? string.Empty;

                var value15 = row.Cell(15).GetValue<string>();
                tipo.IdAteco2007 = !string.IsNullOrEmpty(value15) ? (int?)int.Parse(value15) : null;

                var value16 = row.Cell(16).GetValue<string>();
                tipo.IdFormaGiuridica = !string.IsNullOrEmpty(value16) ? (int?)int.Parse(value16) : null;

                tipo.CodiceREA = row.Cell(17).GetValue<string>() ?? string.Empty;

                var value18 = row.Cell(18).GetValue<string>();
                tipo.NumeroSoci = !string.IsNullOrEmpty(value18) ? (int?)int.Parse(value18) : null;

                var value19 = row.Cell(19).GetValue<string>();
                tipo.NumeroDipendenti = !string.IsNullOrEmpty(value19) ? (int?)int.Parse(value19) : null;

                var value20 = row.Cell(20).GetValue<string>();
                tipo.NumeroCollaboratori = !string.IsNullOrEmpty(value20) ? (int?)int.Parse(value20) : null;

                var value21 = row.Cell(21).GetValue<string>();
                tipo.NumeroIscrittiLibroSoci = !string.IsNullOrEmpty(value21) ? (int?)int.Parse(value21) : null;

                var value22 = row.Cell(22).GetValue<string>();
                tipo.CapitaleSociale = !string.IsNullOrEmpty(value22) ? (decimal?)decimal.Parse(value22) : null;

                var value23 = row.Cell(23).GetValue<string>();
                tipo.DataCapitaleSociale = !string.IsNullOrEmpty(value23) ? (DateTime?)DateTime.Parse(value23) : null;

                tipo.CodiceIscrizioneCCIAA = row.Cell(24).GetValue<string>() ?? string.Empty;

                var value25 = row.Cell(25).GetValue<string>();
                tipo.DataIscrizioneCCIAA = !string.IsNullOrEmpty(value25) ? (DateTime?)DateTime.Parse(value25) : null;

                tipo.SitoWeb = row.Cell(26).GetValue<string>() ?? string.Empty;

                tipo.Iban = row.Cell(27).GetValue<string>() ?? string.Empty;

                tipo.Email = row.Cell(28).GetValue<string>() ?? string.Empty;

                tipo.EmailPEC = row.Cell(29).GetValue<string>() ?? string.Empty;

                var value30 = row.Cell(30).GetValue<string>();
                tipo.DataCostituzione = !string.IsNullOrEmpty(value30) ? (DateTime?)DateTime.Parse(value30) : null;

                var value31 = row.Cell(31).GetValue<string>();
                tipo.DataCessazione = !string.IsNullOrEmpty(value31) ? (DateTime?)DateTime.Parse(value31) : null;

                tipo.Telefono = row.Cell(32).GetValue<string>() ?? string.Empty;

                tipo.Fax = row.Cell(33).GetValue<string>() ?? string.Empty;

                var value34 = row.Cell(34).GetValue<string>();
                tipo.IdComuneSedeLegale = !string.IsNullOrEmpty(value34) ? (int?)int.Parse(value34) : null;

                tipo.IndirizzoSedeLegale = row.Cell(35).GetValue<string>() ?? string.Empty;

                tipo.CivicoSedeLegale = row.Cell(36).GetValue<string>() ?? string.Empty;

                tipo.CapSedeLegale = row.Cell(37).GetValue<string>() ?? string.Empty;

                var value38 = row.Cell(38).GetValue<string>();
                tipo.FlagPrivacy = !string.IsNullOrEmpty(value38) ? (bool?)bool.Parse(value38) : null;

                var value39 = row.Cell(39).GetValue<string>();
                tipo.IdCittadinanza = !string.IsNullOrEmpty(value39) ? (int?)int.Parse(value39) : null;

                var value40 = row.Cell(40).GetValue<string>();
                tipo.DataInizioValidita = !string.IsNullOrEmpty(value40) ? (DateTime?)DateTime.Parse(value40) : null;

                var value41 = row.Cell(41).GetValue<string>();
                tipo.DataFineValidita = !string.IsNullOrEmpty(value41) ? (DateTime?)DateTime.Parse(value41) : null;

                tipo.Utente = row.Cell(42).GetValue<string>() ?? string.Empty;

                tipo.CodiceSezione = row.Cell(43).GetValue<string>() ?? string.Empty;

                tipo.DescrizioneSezione = row.Cell(44).GetValue<string>() ?? string.Empty;

                tipo.DescrizioneComuneSedeLegale = row.Cell(45).GetValue<string>() ?? string.Empty;

                tipo.StatoIscrizione = row.Cell(46).GetValue<string>() ?? string.Empty;

                var value47 = row.Cell(47).GetValue<string>();
                tipo.IdSezioneAlbo = !string.IsNullOrEmpty(value47) ? int.Parse(value47) : 0;

                var value48 = row.Cell(48).GetValue<string>();
                tipo.IdAlboIntermediario = !string.IsNullOrEmpty(value48) ? int.Parse(value48) : 0;

                var value49 = row.Cell(49).GetValue<string>();
                tipo.IdAlboIntermediarioSezione = !string.IsNullOrEmpty(value49) ? int.Parse(value49) : 0;

                var value50 = row.Cell(50).GetValue<string>();
                tipo.IdTipologiaAutorizzazioneIntermediarioSezione = !string.IsNullOrEmpty(value50) ? int.Parse(value50) : 0;

                tipo.Politica = row.Cell(51).GetValue<string>() ?? string.Empty;

                tipo.Recapiti = row.Cell(52).GetValue<string>() ?? string.Empty;

                tipo.DesDenominazioneSede = row.Cell(53).GetValue<string>() ?? string.Empty;

                tipo.IndirizzoSede = row.Cell(54).GetValue<string>() ?? string.Empty;

                tipo.Sede = row.Cell(55).GetValue<string>() ?? string.Empty;

                list.Add(tipo);
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

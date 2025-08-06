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
        public void SaveToExcel(List<ApiItem> items, string filePath)
        {
            List<ApiItem> previousItems = null;
            string previousFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + "_prev.xlsx");

            // Load previous data if it exists
            if (File.Exists(previousFilePath))
            {
                previousItems = CaricaDaExcel(previousFilePath);
            }

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
                            worksheet.Cell(row + 2, col + 1).Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                }

                // Add Changes sheet
                var changesSheet = workbook.Worksheets.Add("Changes");
                changesSheet.Cell(1, 1).Value = "Id";
                changesSheet.Cell(1, 2).Value = "ChangeType";
                changesSheet.Cell(1, 3).Value = "Details";

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

                    int changeRow = 2;
                    foreach (var item in newItems)
                    {
                        changesSheet.Cell(changeRow, 1).Value = item.Id;
                        changesSheet.Cell(changeRow, 2).Value = "New";
                        changesSheet.Cell(changeRow, 3).Value = "New record added";
                        changeRow++;
                    }

                    foreach (var item in missingItems)
                    {
                        changesSheet.Cell(changeRow, 1).Value = item.Id;
                        changesSheet.Cell(changeRow, 2).Value = "Missing";
                        changesSheet.Cell(changeRow, 3).Value = "Record removed";
                        changeRow++;
                    }

                    foreach (var pair in modifiedItems)
                    {
                        changesSheet.Cell(changeRow, 1).Value = pair.Current.Id;
                        changesSheet.Cell(changeRow, 2).Value = "Modified";
                        changesSheet.Cell(changeRow, 3).Value = GetChangeDetails(pair.Previous, pair.Current);
                        changeRow++;
                    }
                }
                else
                {
                    changesSheet.Cell(2, 2).Value = "No previous data for comparison";
                }

                // Save the current file
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

        //
        //public List<ApiItem> CaricaDaExcel(string filePath)
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    var package = new ExcelPackage(new FileInfo(filePath));
        //    var worksheet = package.Workbook.Worksheets[0];

        //    var list = new List<ApiItem>();
        //    var rows = worksheet.Dimension.End.Row;
        //    for (int row = 2; row <= rows; row++) // Assume prima riga = intestazioni
        //    {
        //        var tipo = new ApiItem
        //        {
        //            Id = int.Parse(worksheet.Cells[row, 1].Text),
        //            AzioneView = bool.Parse(worksheet.Cells[row, 2].Text),
        //            AzioneEdit = bool.Parse(worksheet.Cells[row, 3].Text),
        //        };

        //        list.Add(tipo);
        //    }

        //    return list;
        //}

        public List<ApiItem> CaricaDaExcel(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);

            var list = new List<ApiItem>();

            foreach (var row in worksheet.RowsUsed().Skip(1)) // Salta header
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
                    IdTipologiaPersonaGiuridica = int.Parse(row.Cell(9).GetValue<string>()),
                    IdStatoSedePatronato = int.Parse(row.Cell(10).GetValue<string>()),
                    IdSedePatronato = int.Parse(row.Cell(11).GetValue<string>()),
                    CodiceFiscale = row.Cell(12).GetValue<string>(),
                    PartitaIVA = row.Cell(13).GetValue<string>(),
                    RagioneSociale = row.Cell(14).GetValue<string>(),
                    IdAteco2007 = int.Parse(row.Cell(15).GetValue<string>()),
                    IdFormaGiuridica = int.Parse(row.Cell(16).GetValue<string>()),
                    CodiceREA = row.Cell(17).GetValue<string>(),
                    NumeroSoci = int.Parse(row.Cell(18).GetValue<string>()),
                    NumeroDipendenti = int.Parse(row.Cell(19).GetValue<string>()),
                    NumeroCollaboratori = int.Parse(row.Cell(20).GetValue<string>()),
                    NumeroIscrittiLibroSoci = int.Parse(row.Cell(21).GetValue<string>()),
                    CapitaleSociale = decimal.Parse(row.Cell(22).GetValue<string>()),
                    DataCapitaleSociale = DateTime.Parse(row.Cell(23).GetValue<string>()),
                    CodiceIscrizioneCCIAA = row.Cell(24).GetValue<string>(),
                    DataIscrizioneCCIAA = DateTime.Parse(row.Cell(25).GetValue<string>()),
                    SitoWeb = row.Cell(26).GetValue<string>(),
                    Iban = row.Cell(27).GetValue<string>(),
                    Email = row.Cell(28).GetValue<string>(),
                    EmailPEC = row.Cell(29).GetValue<string>(),
                    DataCostituzione = DateTime.Parse(row.Cell(30).GetValue<string>()),
                    DataCessazione = DateTime.Parse(row.Cell(31).GetValue<string>()),
                    Telefono = row.Cell(32).GetValue<string>(),
                    Fax = row.Cell(33).GetValue<string>(),
                    IdComuneSedeLegale = int.Parse(row.Cell(34).GetValue<string>()),
                    IndirizzoSedeLegale = row.Cell(35).GetValue<string>(),
                    CivicoSedeLegale = row.Cell(36).GetValue<string>(),
                    CapSedeLegale = row.Cell(37).GetValue<string>(),
                    FlagPrivacy = bool.Parse(row.Cell(38).GetValue<string>()),
                    IdCittadinanza = int.Parse(row.Cell(39).GetValue<string>()),
                    DataInizioValidita = DateTime.Parse(row.Cell(40).GetValue<string>()),
                    DataFineValidita = DateTime.Parse(row.Cell(41).GetValue<string>()),
                    Utente = row.Cell(42).GetValue<string>(),
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
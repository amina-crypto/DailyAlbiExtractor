using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel; // Corrected typo from ClosedXML.Excel

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
                try
                {
                    previousItems = LoadFromExcel(previousFilePath);
                    Console.WriteLine($"Loaded {previousItems?.Count ?? 0} previous items from {previousFilePath}");
                    if (previousItems != null && previousItems.Any())
                    {
                        Console.WriteLine($"Sample previous CodiceFiscale: {string.Join(", ", previousItems.Take(5).Select(i => i.CodiceFiscale))}");
                    }
                    else
                    {
                        Console.WriteLine("Previous data is empty or invalid, skipping comparison.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading previous file {previousFilePath}: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"No previous file found at {previousFilePath}");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                // Add headers for Data sheet
                var properties = typeof(ApiItem).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = properties[i].Name;
                }

                // Add data rows to Data sheet
                for (int row = 0; row < items.Count; row++)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = properties[col].GetValue(items[row], null);
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

                // Add Changes sheet with same columns as Data
                var changesSheet = workbook.Worksheets.Add("Changes");
                for (int i = 0; i < properties.Length; i++)
                {
                    changesSheet.Cell(1, i + 1).Value = properties[i].Name;
                }

                // Detect changes if previous data exists
                int changeRow = 2;
                bool hasChanges = false;
                if (previousItems != null && previousItems.Any())
                {
                    int previousDataCount = previousItems.Count; // Data rows only
                    int previousTotalCount = previousDataCount + 1; // Include header
                    int currentCount = items.Count;
                    Console.WriteLine($"Previous total count (with header): {previousTotalCount}, Current count: {currentCount}");

                    // Composite key: CodiceFiscale | PartitaIVA
                    var currentDict = items
                        .Where(i => !string.IsNullOrEmpty(i.CodiceFiscale) || !string.IsNullOrEmpty(i.PartitaIVA))
                        .GroupBy(i => (i.CodiceFiscale ?? "") + "|" + (i.PartitaIVA ?? ""))
                        .ToDictionary(g => g.Key, g => g.First()); // Use first item if duplicates

                    var previousDict = previousItems
                        .Where(i => !string.IsNullOrEmpty(i.CodiceFiscale) || !string.IsNullOrEmpty(i.PartitaIVA))
                        .GroupBy(i => (i.CodiceFiscale ?? "") + "|" + (i.PartitaIVA ?? ""))
                        .ToDictionary(g => g.Key, g => g.First()); // Use first item if duplicates

                    var allKeys = currentDict.Keys.Concat(previousDict.Keys).Distinct();

                    foreach (var key in allKeys)
                    {
                        if (!previousDict.ContainsKey(key) && currentDict.ContainsKey(key))
                        {
                            // New item - add full row to Changes sheet
                            var item = currentDict[key];
                            WriteItemToSheet(changesSheet, properties, item, changeRow);
                            hasChanges = true;
                            changeRow++;
                            if (items.Count(g => (g.CodiceFiscale ?? "") + "|" + (g.PartitaIVA ?? "") == key) > 1)
                            {
                                Console.WriteLine($"Duplicate key detected in current data for key: {key}. Using first item.");
                            }
                        }
                        else if (previousDict.ContainsKey(key) && currentDict.ContainsKey(key))
                        {
                            // Check for modifications
                            var prevItem = previousDict[key];
                            var currItem = currentDict[key];
                            if (!AreItemsEqual(prevItem, currItem))
                            {
                                // Modified - add full row of the new version to Changes sheet
                                WriteItemToSheet(changesSheet, properties, currItem, changeRow);
                                hasChanges = true;
                                changeRow++;
                            }
                        }
                    }
                }

                // If no changes, add message
                if (!hasChanges)
                {
                    changesSheet.Cell(2, 1).Value = "No new data";
                }

                // Add length tracking summary (including header for previous count)
                var summaryRow = changeRow + 1;
                changesSheet.Cell(summaryRow, 1).Value = "Previous Record Count (with header)";
                changesSheet.Cell(summaryRow, 2).Value = previousItems?.Count + 1 ?? 0; // Include header
                summaryRow++;
                changesSheet.Cell(summaryRow, 1).Value = "Current Record Count";
                changesSheet.Cell(summaryRow, 2).Value = items.Count;

                // Save the workbook
                workbook.SaveAs(filePath);
                Console.WriteLine($"Excel file saved to: {filePath} with {items.Count} records");

                // Backup after saving to ensure _prev reflects this successful save
                try
                {
                    File.Copy(filePath, previousFilePath, true);
                    Console.WriteLine($"Backed up to: {previousFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error backing up to {previousFilePath}: {ex.Message}");
                }
            }
        }

        private void WriteItemToSheet(IXLWorksheet sheet, PropertyInfo[] properties, ApiItem item, int row)
        {
            for (int col = 0; col < properties.Length; col++)
            {
                var value = properties[col].GetValue(item, null);
                if (value is DateTime dateValue)
                {
                    sheet.Cell(row, col + 1).Value = dateValue;
                    sheet.Cell(row, col + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                }
                else
                {
                    sheet.Cell(row, col + 1).Value = value != null ? value.ToString() : string.Empty;
                }
            }
        }

        private bool AreItemsEqual(ApiItem item1, ApiItem item2)
        {
            var properties = typeof(ApiItem).GetProperties();
            foreach (var prop in properties)
            {
                if (prop.Name == "CodiceFiscale" || prop.Name == "PartitaIVA") continue; // Skip matching keys
                var value1 = prop.GetValue(item1);
                var value2 = prop.GetValue(item2);
                if (value1 != value2 && (value1 == null || !value1.Equals(value2)))
                {
                    return false;
                }
            }
            return true;
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
                foreach (var row in worksheet.RowsUsed())
                {
                    var item = new ApiItem();
                    bool hasValidData = false;
                    foreach (var prop in typeof(ApiItem).GetProperties())
                    {
                        if (headers.TryGetValue(prop.Name, out int col))
                        {
                            var cell = row.Cell(col);
                            if (cell.IsEmpty()) continue; // Skip empty cells
                            try
                            {
                                var valueType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                object value;
                                if (valueType == typeof(string))
                                {
                                    value = cell.GetText() ?? string.Empty; // Use GetText for strings
                                    hasValidData = true;
                                }
                                else if (valueType == typeof(int?) || valueType == typeof(int))
                                {
                                    value = cell.GetValue<int?>(); // Handle nullable int
                                    if (value != null) hasValidData = true;
                                }
                                else if (valueType == typeof(DateTime?) || valueType == typeof(DateTime))
                                {
                                    value = cell.GetValue<DateTime?>(); // Handle nullable DateTime
                                    if (value != null) hasValidData = true;
                                }
                                else if (valueType == typeof(bool?) || valueType == typeof(bool))
                                {
                                    value = cell.GetValue<bool?>(); // Handle nullable bool
                                    if (value != null) hasValidData = true;
                                }
                                else if (valueType == typeof(decimal?) || valueType == typeof(decimal))
                                {
                                    value = cell.GetValue<decimal?>(); // Handle nullable decimal
                                    if (value != null) hasValidData = true;
                                }
                                else if (Utils.IsNumericType(valueType))
                                {
                                    var doubleValue = cell.GetValue<double?>();
                                    value = doubleValue.HasValue ? Convert.ChangeType(doubleValue.Value, valueType) : null;
                                    if (value != null) hasValidData = true;
                                }
                                else
                                {
                                    Console.WriteLine($"Unhandled type {valueType.Name} for property {prop.Name} in row {row.RowNumber()} - Value: {cell.Value}");
                                    value = null; // Skip unhandled types
                                }
                                prop.SetValue(item, value);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Conversion error for property {prop.Name} in row {row.RowNumber()}: {ex.Message} - Value: {cell.Value} - Type: {cell.Value.GetType().Name}");
                            }
                        }
                    }
                    if (hasValidData && row.RowNumber() > 1) items.Add(item); // Exclude header row from items
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
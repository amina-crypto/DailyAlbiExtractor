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
            List<ApiItem> previousItems  = null;


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
                        Console.WriteLine($"Sample previous Ids: {string.Join(", ", previousItems.Take(5).Select(i => i.Id))}");
                    }
                    else
                    {
                        Console.WriteLine("Previous data is empty or invalid, skipping comparison.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading previous file {previousFilePath}: {ex.Message}");
                    previousItems = new List<ApiItem>(); // Fallback to empty list to avoid null reference
                }
            }
            else
            {
                Console.WriteLine($"No previous file found at {previousFilePath}");
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

                if (previousItems != null && previousItems.Any())
                {
                    // Filter out invalid Ids (e.g., <= 0)
                    var validCurrentItems = items.Where(i => i.Id > 0).ToList();
                    var validPreviousItems = previousItems.Where(i => i.Id > 0).ToList();

                    if (!validCurrentItems.Any() || !validPreviousItems.Any())
                    {
                        Console.WriteLine("No valid items for comparison due to invalid Ids.");
                        changesSheet.Cell(2, 2).Value = "No valid data for comparison";
                    }
                    else
                    {
                        // Create dictionaries for efficient lookup
                        var currentDict = validCurrentItems.ToDictionary(i => i.Id);
                        var previousDict = validPreviousItems.ToDictionary(i => i.Id);

                        var allIds = currentDict.Keys.Concat(previousDict.Keys).Distinct();

                        int changeRow = 2;
                        foreach (var id in allIds)
                        {
                            if (!previousDict.ContainsKey(id) && currentDict.ContainsKey(id))
                            {
                                // New item
                                var item = currentDict[id];
                                changesSheet.Cell(changeRow, 1).Value = item.Id;
                                changesSheet.Cell(changeRow, 2).Value = "New";
                                changesSheet.Cell(changeRow, 3).Value = "New record added";
                                changeRow++;
                            }
                            else if (previousDict.ContainsKey(id) && !currentDict.ContainsKey(id))
                            {
                                // Missing item
                                var item = previousDict[id];
                                changesSheet.Cell(changeRow, 1).Value = item.Id;
                                changesSheet.Cell(changeRow, 2).Value = "Missing";
                                changesSheet.Cell(changeRow, 3).Value = "Record removed";
                                changeRow++;
                            }
                            else if (previousDict.ContainsKey(id) && currentDict.ContainsKey(id))
                            {
                                // Check for modifications
                                var prevItem = previousDict[id];
                                var currItem = currentDict[id];
                                if (!AreItemsEqual(prevItem, currItem))
                                {
                                    changesSheet.Cell(changeRow, 1).Value = currItem.Id;
                                    changesSheet.Cell(changeRow, 2).Value = "Modified";
                                    changesSheet.Cell(changeRow, 3).Value = GetChangeDetails(prevItem, currItem);
                                    changeRow++;
                                }
                            }
                        }
                    }
                }
                else
                {
                    changesSheet.Cell(2, 2).Value = "No previous data for comparison";
                }

                // Save the current file with backup
                try
                {
                    // Backup the previous file if it exists and differs from the new path
                    if (File.Exists(filePath) && filePath != previousFilePath)
                    {
                        File.Copy(filePath, previousFilePath, true);
                        Console.WriteLine($"Backed up previous file to {previousFilePath}");
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
                // Handle nulls and compare
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
                                    if (prop.Name == "Id")
                                    {
                                        Console.WriteLine($"Loaded Id: {value}");
                                    }
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
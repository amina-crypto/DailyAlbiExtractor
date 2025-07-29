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

                workbook.SaveAs(filePath);
                Console.WriteLine($"Excel file saved to: {filePath} with {items.Count} records");
            }
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
using ExcelEpplus_api.Interfaces;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Reflection.Metadata.Ecma335;

namespace ExcelEpplus_api.Services
{
    public class ExcelService<T> : IExcelService<T>
    {
        private readonly IConfiguration configuration;
        public ExcelService(IConfiguration configuration)
        {
            this.configuration = configuration;
        }


        public async Task<IActionResult> ReadAsync(string FileName)
        {
            var directoryPath = configuration["Paths:Excel"];
            var filePath = Path.Combine(directoryPath, $"{FileName}.xlsx");
            var items = new List<T>();

            byte[] fileBytes = null;

            int maxRetries = 5;
            int delay = 1000; // 1 second

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;
                        var colCount = worksheet.Dimension.Columns;

                        for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                        {
                            var item = Activator.CreateInstance<T>();
                            var itemType = typeof(T);
                            var properties = itemType.GetProperties();

                            for (int col = 1; col <= colCount; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text;
                                var property = properties[col - 1]; // Assuming the order of properties matches the columns

                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    object convertedValue = null;
                                    if (property.PropertyType == typeof(DateTime))
                                    {
                                        convertedValue = DateTime.FromOADate(double.Parse(cellValue));
                                    }
                                    else
                                    {
                                        convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
                                    }
                                    property.SetValue(item, convertedValue);
                                }
                            }

                            items.Add(item);
                        }
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);
                            fileBytes = memoryStream.ToArray();
                        }
                        break;
                    }
                }
                catch (IOException ex) when (attempt < maxRetries)
                {
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying in {delay / 1000} seconds...");
                    await Task.Delay(delay);
                }
            }

            return new FileContentResult(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = $"{FileName}.xlsx"
            };
        }


        public async Task WriteAsync(string FileName, T request)
        {
            var directoryPath = configuration["Paths:Excel"];
            var filePath = Path.Combine(directoryPath, $"{FileName}.xlsx");

            var properties = typeof(T).GetProperties();

            if (File.Exists(filePath))
            {
                Console.WriteLine($"File '{FileName}.xlsx' already exists. Updating existing file.");

                using (var existingPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = existingPackage.Workbook.Worksheets.FirstOrDefault() ?? existingPackage.Workbook.Worksheets.Add("Sheet1");

                    var rowCount = worksheet.Dimension?.Rows ?? 0;
                    var nextRow = rowCount + 1;

                    for (int col = 0; col < properties.Length; col++)
                    {
                        var propertyValue = properties[col].GetValue(request);
                        worksheet.Cells[nextRow, col + 1].Value = propertyValue;
                    }

                    await existingPackage.SaveAsync();
                }
            }
            else
            {
                Console.WriteLine($"File '{FileName}.xlsx' does not exist. Creating new file.");

                Directory.CreateDirectory(directoryPath);

                using (var newPackage = new ExcelPackage())
                {
                    var worksheet = newPackage.Workbook.Worksheets.Add("Sheet1");

                    // Write header
                    for (int col = 0; col < properties.Length; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = properties[col].Name;
                    }

                    // Write data
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var propertyValue = properties[col].GetValue(request);
                        worksheet.Cells[2, col + 1].Value = propertyValue;
                    }

                    await newPackage.SaveAsAsync(new FileInfo(filePath));
                }
            }

            Console.WriteLine($"Excel file '{FileName}.xlsx' created/updated successfully!");
        }
    }
}


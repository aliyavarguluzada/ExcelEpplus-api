using ExcelEpplus_api.Interfaces;
using OfficeOpenXml;

namespace ExcelEpplus_api.Services
{
    public class ExcelService<T> : IExcelService<T>
    {
        private readonly IConfiguration configuration;
        public ExcelService(IConfiguration configuration)
        {
            this.configuration = configuration;
        }


        //public async Task<byte[]> WriteAsync(string FileName, T request)
        //{
        //    if (string.IsNullOrEmpty(FileName) || request == null)
        //    {
        //        throw new ArgumentException("Invalid file name or request object.");
        //    }

        //    // Create a new Excel package
        //    using (var package = new ExcelPackage())
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        //        // Extract properties of the request object
        //        var properties = typeof(T).GetProperties();

        //        // Write header row
        //        for (int col = 0; col < properties.Length; col++)
        //        {
        //            worksheet.Cells[1, col + 1].Value = properties[col].Name;
        //        }

        //        // Write data rows
        //        worksheet.Cells[2, 1].LoadFromCollection(new List<T> { request }, false);

        //        // Save the Excel package to a memory stream
        //        var memoryStream = new MemoryStream();
        //        await package.SaveAsAsync(memoryStream);

        //        // Set position to the beginning of the stream
        //        memoryStream.Position = 0;

        //        return memoryStream.ToArray();
        //    }


        //public async Task<IActionResult> ReadAsync(IFormFile file)
        //{
        //    var directoryPath = configuration["Paths:Excel"];
        //    //var filePath = Path.Combine(directoryPath, $"{FileName}.xlsx");
        //    var items = new List<T>();

        //    byte[] fileBytes = null;

        //    int maxRetries = 5;
        //    int delay = 1000; // 1 second

        //    for (int attempt = 1; attempt <= maxRetries; attempt++)
        //    {
        //        try
        //        {
        //            using (var package = new ExcelPackage(new FileInfo(filePath)))
        //            {
        //                var worksheet = package.Workbook.Worksheets[0];
        //                var rowCount = worksheet.Dimension.Rows;
        //                var colCount = worksheet.Dimension.Columns;

        //                for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
        //                {
        //                    var item = Activator.CreateInstance<T>();
        //                    var itemType = typeof(T);
        //                    var properties = itemType.GetProperties();

        //                    for (int col = 1; col <= colCount; col++)
        //                    {
        //                        var cellValue = worksheet.Cells[row, col].Text;
        //                        var property = properties[col - 1]; // Assuming the order of properties matches the columns

        //                        if (!string.IsNullOrEmpty(cellValue))
        //                        {
        //                            object convertedValue = null;
        //                            if (property.PropertyType == typeof(DateTime))
        //                            {
        //                                convertedValue = DateTime.FromOADate(double.Parse(cellValue));
        //                            }
        //                            else
        //                            {
        //                                convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
        //                            }
        //                            property.SetValue(item, convertedValue);
        //                        }
        //                    }

        //                    items.Add(item);
        //                }
        //                using (MemoryStream memoryStream = new MemoryStream())
        //                {
        //                    package.SaveAs(memoryStream);
        //                    fileBytes = memoryStream.ToArray();
        //                }
        //                break;
        //            }
        //        }
        //        catch (IOException ex) when (attempt < maxRetries)
        //        {
        //            Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying in {delay / 1000} seconds...");
        //            await Task.Delay(delay);
        //        }
        //    }


        public async Task<byte[]> WriteAsync(string FileName, IEnumerable<T> requests)
        {
            if (string.IsNullOrEmpty(FileName) || requests == null || !requests.Any())
            {
                throw new ArgumentException("Invalid file name or empty request collection.");
            }

            // Create a new Excel package
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Extract properties of the request object (assuming requests are of the same type)
                var properties = typeof(T).GetProperties();

                // Write header row
                for (int col = 0; col < properties.Length; col++)
                {
                    worksheet.Cells[1, col + 1].Value = properties[col].Name;
                }

                // Write data rows
                int row = 2;
                foreach (var request in requests)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var value = properties[col].GetValue(request);
                        worksheet.Cells[row, col + 1].Value = value;
                    }
                    row++;
                }

                // Save the Excel package to a memory stream
                var memoryStream = new MemoryStream();
                await package.SaveAsAsync(memoryStream);

                // Set position to the beginning of the stream
                memoryStream.Position = 0;

                return memoryStream.ToArray();
            }
        }



        public async Task<List<T>> ReadAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("File is empty");
            }

            var items = new List<T>();
            byte[] fileBytes = null;

            int maxRetries = 5;
            int delay = 1000; // 1 second

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var stream = file.OpenReadStream())
                    {
                        using (var package = new ExcelPackage(stream))
                        {
                            var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                            if (worksheet == null)
                            {
                                throw new ArgumentException("Excel file does not contain any worksheet");
                            }

                            var rowCount = worksheet.Dimension.Rows;
                            var colCount = worksheet.Dimension.Columns;

                            for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                            {
                                var item = Activator.CreateInstance<T>();
                                var itemType = typeof(T);
                                var properties = itemType.GetProperties();

                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cellValue = worksheet.Cells[row, col].Value?.ToString();
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

                         
                        }
                    }

                    // Return parsed items and the Excel file bytes
                    return (items);
                }
                catch (IOException ex) when (attempt < maxRetries)
                {
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying in {delay / 1000} seconds...");
                    await Task.Delay(delay);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to process Excel file: {ex.Message}");
                }
            }

            throw new Exception($"Failed to process Excel file after {maxRetries} attempts");
        }




    }


}



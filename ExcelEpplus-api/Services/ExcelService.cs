using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;
using ExcelEpplus_api.Interfaces;
using OfficeOpenXml;

namespace ExcelEpplus_api.Services
{
    public class ExcelService : IExcelService
    {
        private readonly IConfiguration configuration;
        public ExcelService(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        public async Task<List<Employee>> ReadAsync(string FileName)
        {
            var directoryPath = configuration["Paths:Excel"]; 
            var filePath = Path.Combine(directoryPath, $"{FileName}.xlsx");
            var employees = new List<Employee>();

            int maxRetries = 5;
            int delay = 1000; 

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var idCell = worksheet.Cells[row, 1];
                            var nameCell = worksheet.Cells[row, 2];
                            var ageCell = worksheet.Cells[row, 3];

                            if (string.IsNullOrEmpty(idCell.Text))
                            {
                                break;
                            }

                            var employee = new Employee
                            {
                                Id = int.Parse(idCell.Text),
                                Name = nameCell.Text,
                                Age = int.Parse(ageCell.Text)
                            };

                            employees.Add(employee);
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

            return employees;
        }

        public async Task WriteAsync(string FileName, EmployeeRequest request)
        {
            var directoryPath = configuration["Paths:Excel"]; 
            var filePath = Path.Combine(directoryPath, $"{FileName}.xlsx");

            if (File.Exists(filePath))
            {
                Console.WriteLine($"File '{FileName}.xlsx' already exists. Updating existing file.");

                using (var existingPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = existingPackage.Workbook.Worksheets.FirstOrDefault() ?? existingPackage.Workbook.Worksheets.Add("Sheet1");

                    var rowCount = worksheet.Dimension?.Rows ?? 0;
                    var nextRow = rowCount + 1;

                    worksheet.Cells[nextRow, 1].Value = request.Id;
                    worksheet.Cells[nextRow, 2].Value = request.Name;
                    worksheet.Cells[nextRow, 3].Value = request.Age;

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

                    worksheet.Cells[1, 1].Value = request.Id;
                    worksheet.Cells[1, 2].Value = request.Name;
                    worksheet.Cells[1, 3].Value = request.Age;

                    await newPackage.SaveAsAsync(new FileInfo(filePath));
                }
            }

            Console.WriteLine($"Excel file '{FileName}.xlsx' created/updated successfully!");
        }



    }
}


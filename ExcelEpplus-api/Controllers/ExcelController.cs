using ExcelEpplus_api.Entities;
using ExcelEpplus_api.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExcelEpplus_api.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService<Employee> _excelService;

        public ExcelController(IExcelService<Employee> excelService)
        {
            _excelService = excelService;
        }


        [HttpPost("add")]
        public async Task<IActionResult> WriteAsync(string FileName,[FromBody] IEnumerable<Employee> requests)
        {
            try
            {
                var result = await _excelService.WriteAsync(FileName, requests);
                return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{FileName}.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Failed to write Excel file: {ex.Message}");
            }
        }



        [HttpPost("read")]
        public async Task<List<Employee>> ReadAsync(IFormFile file)
        {
            try
            {
                // Call ReadAsync method of IExcelService to read data from the uploaded file
                var items = await _excelService.ReadAsync(file);

                // Return the parsed data as JSON objects
                return items;
            }
            catch (Exception)
            {
                return new List<Employee>();
            }
        }


    }


}

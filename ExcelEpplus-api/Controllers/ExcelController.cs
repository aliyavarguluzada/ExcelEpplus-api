using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;
using ExcelEpplus_api.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExcelEpplus_api.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public ExcelController(IExcelService excelService)
        {
            _excelService = excelService;
        }


        [HttpPost("add")]
        public async Task Write(string FileName, [FromBody] EmployeeRequest request) => await _excelService.WriteAsync(FileName, request);
        [HttpGet("get")]
        public async Task<List<Employee>> Read(string FileName) => await _excelService.ReadAsync(FileName);

    }
}

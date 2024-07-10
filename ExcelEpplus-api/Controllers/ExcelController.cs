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
        public async Task Write(string FileName, [FromBody] Employee request) => await _excelService.WriteAsync(FileName, request);
        [HttpGet("get")]
        public async Task<IActionResult> Read(string FileName) => await _excelService.ReadAsync(FileName);



    }


}

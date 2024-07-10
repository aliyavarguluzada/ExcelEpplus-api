using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;
using Microsoft.AspNetCore.Mvc;

namespace ExcelEpplus_api.Interfaces
{
    public interface IExcelService<T>
    {
        public Task WriteAsync(string FileName, T request);
        public Task<IActionResult> ReadAsync(string FileName);
    }
}

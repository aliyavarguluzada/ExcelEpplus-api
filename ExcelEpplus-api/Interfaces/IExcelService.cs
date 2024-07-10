using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;
using Microsoft.AspNetCore.Mvc;

namespace ExcelEpplus_api.Interfaces
{
    public interface IExcelService<T>
    {
        public Task<byte[]> WriteAsync(string FileName, IEnumerable<T> request);
        public Task<List<T>> ReadAsync(IFormFile file);
    }
}

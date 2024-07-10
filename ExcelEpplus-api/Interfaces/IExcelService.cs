using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;

namespace ExcelEpplus_api.Interfaces
{
    public interface IExcelService<T>
    {
        public Task WriteAsync(string FileName, T request);
        public Task<List<T>> ReadAsync(string FileName);
    }
}

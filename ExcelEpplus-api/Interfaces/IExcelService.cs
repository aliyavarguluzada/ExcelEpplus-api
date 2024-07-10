using ExcelEpplus_api.Core;
using ExcelEpplus_api.Entities;

namespace ExcelEpplus_api.Interfaces
{
    public interface IExcelService
    {
        public Task WriteAsync(string FileName,EmployeeRequest request);
        public Task<List<Employee>> ReadAsync(string FileName);
    }
}

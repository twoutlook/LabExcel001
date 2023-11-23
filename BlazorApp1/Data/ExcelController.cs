using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace BlazorApp1.Data
{
    [ApiController]
    [Route("[controller]")] // This will set the route to /excel
    public class ExcelController:ControllerBase
    {
        [HttpGet("GetExcel")]
        public async Task<IActionResult> GetExcel()
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(new List<ExcelModel>
                {
                    new ExcelModel { Id = 1, Name = "John", Surname = "Doe" },
                    new ExcelModel { Id = 2, Name = "Jane", Surname = "Doe" },
                    new ExcelModel { Id = 3, Name = "John", Surname = "Smith" }
                }, true);
                package.Save();
            }
            stream.Position = 0;
            string excelName = $"Excel_{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }   
    }
}

using Microsoft.AspNetCore.Mvc;
using NPOI.Extensions.Web;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExportApiDemo.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExportController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<ExportController> _logger;

        public ExportController(ILogger<ExportController> logger)
        {
            _logger = logger;
        }

        XSSFWorkbook GenerateXSSFExcel()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    if (j == 14)
                    {
                        row.CreateCell(j).SetCellFormula("B2");
                        continue; 
                    }
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
            var dt = sheet1.ToDataTable(true, true);
            return workbook;
        }

        HSSFWorkbook GenerateHSSFExcel()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is My Sample");
            var dt=sheet1.ToDataTable(false);
            return workbook;
        }

        [HttpGet]
        public ActionResult Get()
        {
            return new XSSFFileResult(GenerateXSSFExcel(), "test1.xlsx");
        }
    }
}

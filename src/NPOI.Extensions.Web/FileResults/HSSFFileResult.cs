using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Threading.Tasks;

namespace NPOI.Extensions.Web
{
    public class HSSFFileResult : FileStreamResult
    {
        private HSSFWorkbook _workbook;
        public HSSFFileResult(HSSFWorkbook workbook, string fileDownloadName) : base(new MemoryStream(1), "application/vnd.ms-excel")
        {
            FileDownloadName = fileDownloadName;
            _workbook = workbook;

            FileStream = ConvertWorkbookToStream(_workbook);
        }
        private Stream ConvertWorkbookToStream(HSSFWorkbook workbook)
        {
            MemoryStream ms = RecyclableMemory.GetStream();
            workbook.Write(ms);
            ms.Position = 0;
            return ms;            
        }
    }
}

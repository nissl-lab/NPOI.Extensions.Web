using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Threading.Tasks;

namespace NPOI.Extensions.Web
{
    public class XSSFFileResult : FileStreamResult
    {
        private XSSFWorkbook _workbook;
        public XSSFFileResult(XSSFWorkbook workbook, string fileDownloadName) : base(new MemoryStream(1), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = fileDownloadName;
            _workbook = workbook;

            FileStream = ConvertWorkbookToStream(_workbook);
        }
        private Stream ConvertWorkbookToStream(XSSFWorkbook workbook)
        {
            MemoryStream ms = RecyclableMemory.GetStream();
            workbook.Write(ms, true);
            ms.Position=0;
            return ms;
        }
    }
    



}

using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.Streaming;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Extensions.Web
{
    public class SXSSFFileResult : FileStreamResult
    {
        private SXSSFWorkbook _workbook;
        public SXSSFFileResult(SXSSFWorkbook workbook, string fileDownloadName) : base(new MemoryStream(1), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = fileDownloadName;
            _workbook = workbook;

            FileStream = ConvertWorkbookToStream(_workbook);
        }
        private Stream ConvertWorkbookToStream(SXSSFWorkbook workbook)
        {
            MemoryStream ms = RecyclableMemory.GetStream();
            workbook.Write(ms);
            ms.Position = 0;
            return ms;
        }
    }
}

using Microsoft.AspNetCore.Mvc;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Extensions.Web
{
    public class XWPFFileResult: FileStreamResult
    {
        XWPFDocument _document;
        public XWPFFileResult(XWPFDocument doc, string fileDownloadName) : base(new MemoryStream(1), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        {
            FileDownloadName = fileDownloadName;
            _document = doc;
            this.FileStream = ConvertDocumentToStream(doc);
        }
        private Stream ConvertDocumentToStream(XWPFDocument doc)
        {
            MemoryStream ms = RecyclableMemory.GetStream();
            doc.Write(ms);
            ms.Position = 0;
            return ms;
        }
        public override Task ExecuteResultAsync(ActionContext context)
        {
            return base.ExecuteResultAsync(context);
        }
    }
}

# NPOI.Extensions.Web

### FileResult
- XSSFFileResult
- SXSSFFileResult
- HSSFFileResult
- XWPFFileResult

*Sample Code*
```
[HttpGet]
public ActionResult Get()
{
    XSSFWorkbook workbook = new XSSFWorkbook();
    ISheet sheet1 = workbook.CreateSheet("Sheet1");
    return new XSSFFileResult(workbook, "test1.xlsx");
}
```

### DataSet/DataTable Support
- IWorkbook.ToDataSet()
- ISheet.ToDataTable()

*Sample Code*

```var dt=sheet1.ToDataTable(false);```

# Planning
### TempStorage Support
- Azure Bloc
- AWS S3
- Local File



using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace NPOI.Extensions.Web
{
    public static class DataTableExtension
    {
        public static DataSet ToDataSet(this IWorkbook workbook, bool firstRowAsHeader, bool showCalculatedFormulaValue = false)
        {
            DataSet ds = new DataSet();
            for (int i = 0; i < workbook.NumberOfSheets;i++)
            {
                var sheet=workbook.GetSheetAt(0);
                var dt=sheet.ToDataTable(firstRowAsHeader, showCalculatedFormulaValue);
                ds.Tables.Add(dt);
            }
            return ds;
        }
        public static DataTable ToDataTable(this ISheet sheet, bool firstRowAsHeader, bool showCalculatedFormulaValue=false)
        {
            DataTable dt = new DataTable(sheet.SheetName);

            HSSFDataFormatter formatter = new HSSFDataFormatter();
            var formulaEvalator=sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row=sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                if (i == 0) //don't support merged cell so far
                {
                    if (firstRowAsHeader)
                    {
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            var c = row.GetCell(j);
                            if (c == null)
                            {
                                dt.Columns.Add("");
                            }
                            else
                            {
                                dt.Columns.Add(formatter.FormatCellValue(c));
                            }
                        }
                        continue;
                    }
                }
                
                DataRow dr = null;
                for (int j = 0; j < row.LastCellNum; j++)
                {                    
                    if (j == 0)
                    {
                        for (int k = dt.Columns.Count; k < row.LastCellNum; k++)
                        {
                            dt.Columns.Add("Column " + k);
                        }
                        dr = dt.NewRow();
                    }
                    var c = row.GetCell(j);
                    if (c == null)
                    {
                        dr[j] = "";
                        continue;
                    }
                    switch (c.CellType) 
                    {
                        case CellType.Numeric:
                            dr[j] =  c.NumericCellValue;
                            break;
                        case CellType.String:
                            dr[j] = formatter.FormatCellValue(c);
                            break;
                        case CellType.Blank:
                            dr[j] = "";
                            break;
                        case CellType.Boolean:
                            dr[j] = c.BooleanCellValue;
                            break;
                        case CellType.Formula:
                            if (showCalculatedFormulaValue)
                            {
                                var cellvalue = formulaEvalator.Evaluate(c);
                                switch (cellvalue.CellType)
                                {
                                    case CellType.Numeric:
                                        dr[j] = cellvalue.NumberValue;
                                        break;
                                    case CellType.String:
                                        dr[j] = cellvalue.StringValue;
                                        break;
                                    case CellType.Boolean:
                                        dr[j] = cellvalue.BooleanValue;
                                        break;
                                    case CellType.Error:
                                        if(c.ErrorCellValue== FormulaError.NULL.Code)
                                            dr[j] = "#NULL!";
                                        else if(c.ErrorCellValue == FormulaError.DIV0.Code)
                                            dr[j] = "#DIV/0!";
                                        else if (c.ErrorCellValue == FormulaError.NUM.Code)
                                            dr[j] = "#NUM!";
                                        else if (c.ErrorCellValue == FormulaError.NAME.Code)
                                            dr[j] = "#NAME!";
                                        break;
                                    case CellType.Blank:
                                        dr[j] = "";
                                        break;
                                }
                            }
                            else
                                dr[j] = c.CellFormula;
                            break;
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}

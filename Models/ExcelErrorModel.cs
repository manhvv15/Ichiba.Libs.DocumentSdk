namespace Ichiba.Libs.DocumentSdk.Models;

public record ExcelErrorModel(string FieldName, string Sheetname, string CellName, string ErrorMessage);

using Ichiba.Libs.DocumentSdk.Abstractions;

namespace Ichiba.Libs.DocumentSdk.Models;

public class ImportExcelResponse<T> where T : DocumentItemBase, new()
{
    public ImportExcelResponse(bool isSuccess, IEnumerable<T> response, string sheetName, int lastCol)
    {
        IsSuccess = isSuccess;
        Data = response;
        SheetName = sheetName;
        LastCol = lastCol;
    }
    public ImportExcelResponse(string errorMessage)
    {
        IsSuccess = false;
        ErrorMessage = errorMessage;
    }

    public ImportExcelResponse(bool isSuccess, IEnumerable<T>? data, string? errorMessage, string sheetName, int lastCol)
    {
        IsSuccess = isSuccess;
        Data = data;
        ErrorMessage = errorMessage;
        SheetName = sheetName;
        LastCol = lastCol;
    }

    public bool IsSuccess { get; set; }
    public IEnumerable<T>? Data { get; set; }
    public string? ErrorMessage { get; set; }
    public string SheetName { get; set; }
    public int LastCol { get; set; }
}

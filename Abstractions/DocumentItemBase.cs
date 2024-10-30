using Ichiba.Libs.DocumentSdk.Extend;
using Ichiba.Libs.DocumentSdk.Helpers;
using Ichiba.Libs.DocumentSdk.Models;

namespace Ichiba.Libs.DocumentSdk.Abstractions;

public abstract class DocumentItemBase
{
    [IgnoreProperty]
    public List<ExcelErrorModel> Errors { get; set; } = new();

    public void AddError(string fieldName, string sheetName, string cellName, string errorMessage)
    {
        Errors ??= new List<ExcelErrorModel>();
        Errors.Add(new ExcelErrorModel(fieldName, sheetName, cellName, errorMessage));
    }

    public void AddError(string filedName, string sheetName, int rowIndex, int colIndex, string errorMessage)
    {

        Errors ??= new List<ExcelErrorModel>();
        var cellName = AsposeHelper.GetCellName(rowIndex, colIndex);
        Errors.Add(new ExcelErrorModel(filedName, sheetName, cellName, errorMessage));
    }
}

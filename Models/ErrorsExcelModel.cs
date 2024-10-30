namespace Ichiba.Libs.DocumentSdk.Models;

public class ErrorsExcelModel<T> : List<Tuple<int, T, string>>
{
    public void Add(int rowIndex, T model, string error)
    {
        this.Add(new Tuple<int, T, string>(rowIndex, model, error));
    }
}

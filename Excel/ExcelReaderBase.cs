using MiniExcelLibs;

namespace Ichiba.Libs.DocumentSdk.Excel;

public abstract class ExcelReaderBase<T> : IDisposable
    where T : class, new()
{
    private readonly Stream _stream;

    public ExcelReaderBase(string path)
    {
        _stream = File.OpenRead(path);
    }

    public ExcelReaderBase(Stream stream)
    {
        _stream = stream;
    }

    public IEnumerable<T> Read(string sheetName,
        ExcelType excelType = ExcelType.UNKNOWN,
        string startCell = "A1"
        , IConfiguration? configuration = null)
    {
        var data = MiniExcel.Query<T>(_stream, sheetName: sheetName, excelType, startCell, configuration);

        Validate(data);

        return data;
    }

    protected abstract void Validate(IEnumerable<T> items);

    public void Dispose()
    {
        _stream.Dispose();
    }
}

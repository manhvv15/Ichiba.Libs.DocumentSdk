using Aspose.Cells;
using Ichiba.Libs.DocumentSdk.Constants;
using Ichiba.Libs.DocumentSdk.Enums;
using Ichiba.Libs.DocumentSdk.Helpers;
using Ichiba.Libs.DocumentSdk.Models;
using MiniExcelLibs;
using LoadOptions = Aspose.Cells.LoadOptions;

namespace Ichiba.Libs.DocumentSdk.Abstractions;

internal abstract class BaseDocument<T>(
    IHttpClientFactory httpClientFactory,
    IEnumerable<IDocumentValidator<T>> validators)
    : IDisposable
    where T : DocumentItemBase, new()
{
    private readonly HttpClient _client = httpClientFactory.CreateClient();

    protected async Task<ImportExcelResponse<T>> ReadFileAsync(Stream stream, ImportExcelRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var data = await ImportExcelByAsposeAsync(stream, request, cancellationToken);
        if (!data.IsSuccess)
        {
            return data;
        }
        if (validators != null && validators.Any())
        {
            foreach (var validator in validators)
            {
                await validator.ValidateAsync(data.Data, data.SheetName, data.LastCol + 1, cancellationToken);
            }
            data = new ImportExcelResponse<T>(!data.Data.Any(x => x.Errors.Any()), data.Data, data.ErrorMessage, data.SheetName, data.LastCol);
        }
        return data;
    }

    protected async Task<ImportExcelResponse<T>> ReadFileAsync(string filePath, ImportExcelRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var fileStream = await GetFileTemplateAsync(filePath, cancellationToken);
        return await ReadFileAsync(fileStream, request, cancellationToken);
    }

    protected async Task<Byte[]> WriteFileAsync(ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var templateData = await GetFileTemplateAsync(request.Uri, cancellationToken);
        return await ExportExcelByAspose(request, templateData);
    }

    protected async Task<Byte[]>? WriteFileErrorAsync(Stream stream, ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var result = await AsposeHelper.AddErrorAsync(stream, request.Errors, request.FileName, cancellationToken);
        return result;
    }

    public async Task<Byte[]> WriteFileAsync(Stream file, ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        return await ExportExcelByAspose(request, file);
    }


    public async Task<Stream> AddProtectSheet(Stream file, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var result = await AsposeHelper.AddProtectedSheetAsync(file, CommonConstants.NameSheetKey, CommonConstants.PasswordSheetKey, CommonConstants.KeyProtected, CommonConstants.CellContainKey, cancellationToken);
        return result;
    }

    private async Task<Stream> GetFileTemplateAsync(string uri, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var templateFileData = await _client.GetByteArrayAsync(uri, cancellationToken);
        if (templateFileData.Length != 0)
        {
            return new MemoryStream(templateFileData);
        }

        throw new ApplicationException("File not found");
    }

    #region mini excel

    private IEnumerable<T> ImportExcelByMiniExcel(Stream stream, string? sheetName = null,
        ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1"
        , IConfiguration? configuration = null)
    {
        var data = stream.Query<T>(sheetName: sheetName, excelType, startCell, configuration);
        return data;
    }

    #endregion mini excel

    #region aspose

    private async Task<ImportExcelResponse<T>> ImportExcelByAsposeAsync(Stream stream, ImportExcelRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var workbook = new Workbook(stream);
        if (request.IsCheckProtectedSheet)
        {
            var resultValidateKey = await AsposeHelper.ValidateKeyAsync(workbook, CommonConstants.NameSheetKey, CommonConstants.PasswordSheetKey, CommonConstants.KeyProtected, CommonConstants.CellContainKey, cancellationToken);
            if (!resultValidateKey)
            {
                return new ImportExcelResponse<T>(ErrorMessageConstants.NotValidKeySheet);
            }
        }
        var validatedData = await AsposeHelper.GetDataSheetAsync<T>(workbook, cancellationToken);
        return validatedData;
    }

    private async Task<Byte[]> ExportExcelByAspose(ExportSingleRequest request, Stream template)
    {
        var loadOptions = new LoadOptions(LoadFormat.Xlsx);
        var wb = new Workbook(template, loadOptions);
        var wd = new WorkbookDesigner(wb);

        AsposeHelper.SetResource(wd, request.Data, request.ColumnGroups is not null);
        await AsposeHelper.SetupImages(request.Images, wd, _client);
        wd.Workbook.CalculateFormula();
        wd.Process();

        if (request.ColumnGroups is not null)
        {
            foreach (var columnItem in request.ColumnGroups)
            {
                if (string.IsNullOrEmpty(columnItem.RangeName) || columnItem.Columns is null)
                {
                    continue;
                }

                AsposeHelper.GroupCell(wb, columnItem.RangeName, columnItem.Columns.ToArray(), columnItem.Type);
            }
        }

        var outputFile = new MemoryStream();
        wb.Save(outputFile, MapToFormat(request.ExportType()));
        var file = outputFile.ToArray();
        var ms = new MemoryStream();
        ms.Write(file, 0, file.Length);
        ms.Position = 0;

        return await Task.FromResult(ms.ToArray());
    }

    #endregion aspose

    private SaveFormat MapToFormat(ExportType exportType)
    {
        switch (exportType)
        {
            case ExportType.Xlsx:
                return SaveFormat.Xlsx;

            case ExportType.Docx:
                return SaveFormat.Docx;

            case ExportType.Pdf:
                return SaveFormat.Pdf;

            default:
                throw new ArgumentOutOfRangeException(nameof(exportType), exportType, null);
        }
    }

    public void Dispose()
    {
        _client.Dispose();
    }
}

using Ichiba.Libs.DocumentSdk.Abstractions;
using Ichiba.Libs.DocumentSdk.Models;

namespace Ichiba.Libs.DocumentSdk.Services;

internal class ExcelService<T>(IHttpClientFactory httpClientFactory, IEnumerable<IDocumentValidator<T>> validators)
    : BaseDocument<T>(httpClientFactory, validators), IExcelService<T>
    where T : DocumentItemBase, new()
{
    public async Task<ImportExcelResponse<T>> ReadAsync(string filePath, ImportExcelRequest request, CancellationToken cancellationToken = default) => await ReadFileAsync(filePath, request, cancellationToken);

    public async Task<ImportExcelResponse<T>> ReadAsync(Stream file, ImportExcelRequest request, CancellationToken cancellationToken = default) => await ReadFileAsync(file, request, cancellationToken);

    public async Task<DocumentResponse> WriteAsync(ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        var document = await WriteFileAsync(request, cancellationToken);
        return new DocumentResponse()
        {
            Success = true,
            FileName = request.FileName,
            FileExtension = request.FileExtension,
            Data = document
        };
    }

    public async Task<DocumentResponse> WriteAsync(Stream file, ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        var document = await WriteFileAsync(request, cancellationToken);
        return new DocumentResponse()
        {
            Success = true,
            FileName = request.FileName,
            FileExtension = request.FileExtension,
            Data = document
        };
    }

    public async Task<DocumentResponse> WriteErrorAsync(Stream file, ExportSingleRequest request, CancellationToken cancellationToken = default)
    {
        var document = await WriteFileErrorAsync(file, request, cancellationToken);
        return new DocumentResponse()
        {
            Success = true,
            FileName = request.FileName,
            FileExtension = request.FileExtension,
            Data = document
        };
    }
}

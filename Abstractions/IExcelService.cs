using Ichiba.Libs.DocumentSdk.Models;

namespace Ichiba.Libs.DocumentSdk.Abstractions;

public interface IExcelService<T> where T : DocumentItemBase, new()
{
    Task<ImportExcelResponse<T>> ReadAsync(string filePath, ImportExcelRequest request, CancellationToken cancellationToken = default);

    Task<ImportExcelResponse<T>> ReadAsync(Stream file, ImportExcelRequest request, CancellationToken cancellationToken = default);

    Task<DocumentResponse> WriteAsync(ExportSingleRequest request, CancellationToken cancellationToken = default);

    Task<DocumentResponse> WriteAsync(Stream file, ExportSingleRequest request, CancellationToken cancellationToken = default);
    Task<DocumentResponse> WriteErrorAsync(Stream file, ExportSingleRequest request, CancellationToken cancellationToken = default);

    Task<Stream> AddProtectSheet(Stream file, CancellationToken cancellationToken = default);
}

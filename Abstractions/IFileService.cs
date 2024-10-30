using Ichiba.Libs.DocumentSdk.Models;

namespace Ichiba.Libs.DocumentSdk.Abstractions;

public interface IFileService
{
    Task<DocumentResponse> WriteAsync(ExportSingleRequest request);
    Task<DocumentResponse> WriteAsync(ExportMultipleRequest request);
}

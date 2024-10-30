using Ichiba.Libs.DocumentSdk.Constants;
using Ichiba.Libs.DocumentSdk.Models;
using RestEase;

namespace Ichiba.Libs.DocumentSdk.Connectors;

/// <summary>
/// Product APIs
/// Lib. documentation https://github.com/canton7/RestEase
/// </summary>
[Header("User-Agent",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36")]
public interface IDocumentConnector
{
    /// <summary>
    /// Export single
    /// </summary>
    /// <param name="body"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [Post(Endpoints.ExportSingleFile)]
    public Task<DocumentResponse> Export([Body] ExportSingleRequest body,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Export Multi
    /// </summary>
    /// <param name="body"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [Post(Endpoints.ExportMultiFile)]
    public Task<DocumentResponse> Exports([Body] ExportMultipleRequest body,
    CancellationToken cancellationToken = default);

    [Post(Endpoints.MergePdfDocument)]
    public Task<MergePdfDocumentsResponse> MergePdfDocuments([Body] MergePdfDocumentsRequest body,
        CancellationToken cancellationToken = default);
}

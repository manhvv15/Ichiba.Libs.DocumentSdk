using Ichiba.Libs.DocumentSdk.Abstractions;
using Ichiba.Libs.DocumentSdk.Connectors;
using Ichiba.Libs.DocumentSdk.Models;

namespace Ichiba.Libs.DocumentSdk.Services;

internal class PdfService(IDocumentConnector documentConnector) : IPdfService
{
    public async Task<DocumentResponse> WriteAsync(ExportSingleRequest request)
    {
        return await documentConnector.Export(request);
    }

    public async Task<DocumentResponse> WriteAsync(ExportMultipleRequest request)
    {
        return await documentConnector.Exports(request);
    }
}

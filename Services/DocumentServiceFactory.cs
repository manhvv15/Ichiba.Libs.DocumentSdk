using Ichiba.Libs.DocumentSdk.Abstractions;
using Ichiba.Libs.DocumentSdk.Enums;
using Microsoft.Extensions.DependencyInjection;

namespace Ichiba.Libs.DocumentSdk.Services;

internal class DocumentServiceFactory(IServiceProvider serviceProvider) : IDocumentServiceFactory
{
    public IExcelService<T> Create<T>(FileType type) where T : DocumentItemBase, new()
    {
        return serviceProvider.GetRequiredService<IExcelService<T>>();
    }

    public IFileService Create(TemplateType type)
    {
        return type switch
        {
            TemplateType.Docx => serviceProvider.GetRequiredService<IWordService>(),
            TemplateType.Pdf => serviceProvider.GetRequiredService<IPdfService>(),
            _ => throw new NotSupportedException()
        };
    }
}

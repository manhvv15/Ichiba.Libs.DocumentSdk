using Ichiba.Libs.DocumentSdk.Enums;

namespace Ichiba.Libs.DocumentSdk.Abstractions;

public interface IDocumentServiceFactory
{
    IExcelService<T> Create<T>(FileType type) where T : DocumentItemBase, new();
    IFileService Create(TemplateType type);
}

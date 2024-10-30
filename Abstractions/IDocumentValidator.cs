namespace Ichiba.Libs.DocumentSdk.Abstractions;

public interface IDocumentValidator<in T> where T : class
{
    Task ValidateAsync(IEnumerable<T> entities, string sheetName, int lastCol, CancellationToken cancellationToken = default);
}

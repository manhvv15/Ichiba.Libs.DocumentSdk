namespace Ichiba.Libs.DocumentSdk.Models;

public class DocumentResponse
{
    public bool Success { get; set; }
    public string Code { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string FileExtension { get; set; } = string.Empty;
    public byte[]? Data { get; set; }
}

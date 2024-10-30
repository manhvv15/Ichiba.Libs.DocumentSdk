namespace Ichiba.Libs.DocumentSdk.Models;

public class BarCodeDetail
{
    public string Value { get; set; }
    public int Weight { get; set; }
    public int Height { get; set; }
    public bool IsQrCode { get; set; }
    public bool DisplayValueBarCode { get; set; } = true; // Có hiển thị value của barcode không
}

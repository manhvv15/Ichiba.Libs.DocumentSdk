using Ichiba.Libs.DocumentSdk.Enums;

namespace Ichiba.Libs.DocumentSdk.Models;

public class ExportSingleRequest
{
    /// <summary>
    /// Loại file
    /// </summary>
    public string FileType { get; set; } = string.Empty;

    /// <summary>
    /// Loại file mong muốn nhận về
    /// </summary>
    public string FileExtension { get; set; } = string.Empty;

    /// <summary>
    /// Tên file mong muốn trả về
    /// </summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// Template url lấy từ storage Service
    /// </summary>
    public string Uri { get; set; } = string.Empty;

    /// <summary>
    /// danh sách image logo,cod, truyền vào là 1 link cdn
    /// </summary>
    public Dictionary<string, ImageDetail> Images { get; set; } = new Dictionary<string, ImageDetail>();

    /// <summary>
    /// danh sách barcode
    /// </summary>
    public Dictionary<string, BarCodeDetail> BarCodes { get; set; } = new Dictionary<string, BarCodeDetail>();

    /// <summary>
    /// dữ liệu xuất ra file, dạng json string
    /// </summary>
    public Dictionary<string, string> Data { get; set; } = new Dictionary<string, string>();

    /// <summary>
    ///
    /// </summary>
    public List<GroupColumnItem?> ColumnGroups { get; set; }

    public List<ExcelErrorModel> Errors { get; set; }

    public ExportType ExportType()
    {
        var names = Enum.GetNames(typeof(ExportType));
        if (names.Any(x => x.ToLower().Equals(this.FileExtension.ToLower())))
        {
            return Enum.Parse<ExportType>(FileExtension, true);
        }

        throw new ApplicationException();
    }
    public TemplateType RequestType()
    {
        var names = Enum.GetNames(typeof(TemplateType));
        if (names.Any(x => x.ToLower().Equals(this.FileType.ToLower())))
        {
            return Enum.Parse<TemplateType>(FileType, true);
        }

        throw new ApplicationException();
    }
}

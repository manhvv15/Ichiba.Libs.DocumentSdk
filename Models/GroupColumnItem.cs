using Ichiba.Libs.DocumentSdk.Enums;

namespace Ichiba.Libs.DocumentSdk.Models;

public class GroupColumnItem
{
    public string RangeName { get; set; }
    public List<int> Columns { get; set; }
    public GroupColumnItemTypeEnum Type { get; set; } = GroupColumnItemTypeEnum.Row;
}

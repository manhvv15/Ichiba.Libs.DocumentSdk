using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Globalization;
using System.Reflection;
using Aspose.Cells;
using Ichiba.Libs.DocumentSdk.Abstractions;
using Ichiba.Libs.DocumentSdk.Constants;
using Ichiba.Libs.DocumentSdk.Enums;
using Ichiba.Libs.DocumentSdk.Extend;
using Ichiba.Libs.DocumentSdk.Extensions;
using Ichiba.Libs.DocumentSdk.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace Ichiba.Libs.DocumentSdk.Helpers;

public static class AsposeHelper
{
    #region import

    /// <summary>
    /// Gets the data from the specified worksheet and maps it to a DataTable based on the provided type.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the data.</param>
    /// <param name="type">The type of the object to map the data to.</param>
    /// <returns>The DataTable containing the mapped data.</returns>
    public static async Task<ImportExcelResponse<T>> GetDataSheetAsync<T>(Workbook workbook, CancellationToken cancellationToken = default) where T : DocumentItemBase, new()
    {
        return await Task.Run(async () =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var validateData = new List<T>();
            Worksheet? worksheet = null;
            var firtCol = 0;
            var firtRow = 0;
            var endCol = 0;
            var endRow = 0;
            var headerRow = 0;
            var isValidateData = true;
            var sheetName = "";

            var attributeType = typeof(T).GetCustomAttribute(typeof(WorkSheetAttribute), true);
            if (attributeType != null)
            {
                var workSheetAttribute = (WorkSheetAttribute)attributeType;
                sheetName = workSheetAttribute.GetSheetName();
                worksheet = workbook.Worksheets[sheetName];
                if (worksheet is null)
                {
                    return new ImportExcelResponse<T>(ErrorMessageConstants.WorksheetNotFound);
                }

                headerRow = workSheetAttribute.GetHeaderRow();
                firtCol = workSheetAttribute.GetStartCol();
                firtRow = workSheetAttribute.GetStartRow();
                endCol = workSheetAttribute.GetEndCol() >= 0 && workSheetAttribute.GetEndCol() >= firtCol ? workSheetAttribute.GetEndCol() : worksheet.Cells.MaxDataColumn;
                endRow = workSheetAttribute.GetEndRow() >= 0 && workSheetAttribute.GetEndRow() >= firtRow ? workSheetAttribute.GetEndRow() : worksheet.Cells.MaxDataRow;
                isValidateData = workSheetAttribute.IsValidate();
            }
            else
            {
                sheetName = CommonConstants.DefaultNameSheetGetValue;
                worksheet = workbook.Worksheets[sheetName];
                if (worksheet is null)
                {
                    return new ImportExcelResponse<T>(ErrorMessageConstants.WorksheetNotFound);
                }

                endCol = worksheet.Cells.MaxDataColumn;
                endRow = worksheet.Cells.MaxDataRow;
            }

            //lấy header file
            var headers = new Dictionary<int, string>();
            for (int j = firtCol; j <= endCol; j++)
            {
                var headerCell = worksheet.Cells[headerRow, j];
                headers[j] = headerCell.StringValue;
            };

            // dictionary để lưu trữ giá trị unique
            var uniqueValueTracker = new Dictionary<string, HashSet<object>>();

            for (int i = headerRow + 1; i <= endRow; i++)
            {
                // tạo instance của type
                var instance = new T();

                for (int j = 0; j <= endCol; j++)
                {
                    // lấy ra ô hiện tại
                    var cell = worksheet.Cells[i, j];
                    // lấy ra tên cột tương ứng với ô hiện tại
                    if (!headers.TryGetValue(j, out var headerName))
                    {
                        // không có thì next
                        continue;
                    }

                    var isHeaderRequired = false;
                    if (headerName.Contains("*") || headerName.Contains("(*)"))
                    {
                        headerName = headerName.Replace("(*)", "").Replace("*", "").Replace("()", "").Trim();
                        isHeaderRequired = true;
                    }

                    // lấy ra property tương ứng với tên cột
                    var property = typeof(T).GetProperties().FirstOrDefault(p => p.GetPropertyNameOrAlias().Equals(headerName, StringComparison.OrdinalIgnoreCase));

                    // không có thì next
                    if (property == null)
                    {
                        continue;
                    }

                    // lấy giá trị từ ô hiện tại và set vào property, nếu không convert được sẽ add error
                    await SetValueCellToPropertyAsync<T>(cell, worksheet.Name, instance, property, uniqueValueTracker, isValidateData, isHeaderRequired, cancellationToken);
                }

                //lấy ra value  instance và add vào datatable
                validateData.Add(instance);
            }
            return new ImportExcelResponse<T>(!validateData.Any(x => x.Errors.Any()), validateData, sheetName, endCol);
        }, cancellationToken);
    }

    /// <summary>
    /// Validates the key in the specified worksheet.
    /// </summary>
    /// <param name="workbook">The workbook containing the worksheet.</param>
    /// <param name="sheetName">The name of the worksheet.</param>
    /// <param name="password">The password to unprotect the worksheet.</param>
    /// <param name="key">The key to validate.</param>
    /// <param name="cell">The cell containing the value to compare with the key.</param>
    /// <returns>True if the key is valid, otherwise false.</returns>
    public static async Task<bool> ValidateKeyAsync(Workbook workbook, string sheetName, string password, string key, string cell, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            // lấy sheet theo sheet name
            var worksheetKey = workbook.Worksheets[sheetName];
            //nếu sheet đã được mở hoặc sheet k tồn tại
            if (worksheetKey is null || !worksheetKey.IsProtected)
            {
                return false;
            }

            //mở khóa sheet
            try
            {
                worksheetKey.Unprotect(password);
                //lấy value theo cell
                string cellValue = worksheetKey.Cells[cell].StringValue;
                if (key != cellValue)
                {
                    return false;
                }
                return true;
            }
            catch (Exception) //exception, thường là mật khẩu sai
            {
                return false;
            }
        }, cancellationToken);

    }

    /// <summary>
    /// add error to file
    /// </summary>
    /// <param name="stream">Data file to add error</param>
    /// <param name="errors">The dictionary to store the errors</param>
    /// <param name="fileName">name of file return</param>
    /// <returns></returns>
    public static async Task<Byte[]>? AddErrorAsync(Stream stream, List<ExcelErrorModel> errors, string fileName, CancellationToken cancellationToken = default)
    {
        return await Task.Run(async () =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            if (errors is null || !errors.Any())
            {
                return null;
            }

            var workbook = new Workbook(stream);
            var sheets = errors.Select(x => x.Sheetname).Distinct().ToList();

            foreach (var sheetName in sheets)
            {
                var sheet = workbook.Worksheets[sheetName];
                if (sheet is null)
                    continue;

                foreach (var error in errors.Where(x => x.Sheetname == sheetName).GroupBy(x => x.CellName).ToList())
                {
                    var cell = sheet.Cells[error.Key];
                    StyleErrorCell(cell);
                    var errorMessages = string.Join("\n", error.Select(x => x.ErrorMessage));
                    await AddCommentAsync(cell, errorMessages, cancellationToken);
                }
            }

            var outputFile = new MemoryStream();
            workbook.Save(outputFile, SaveFormat.Xlsx);
            var file = outputFile.ToArray();
            var ms = new MemoryStream();
            ms.Write(file, 0, file.Length);
            ms.Position = 0;

            return ms.ToArray();
        }, cancellationToken);


    }

    /// <summary>
    /// Adds an error to the specified cell and stores the error message in the errors dictionary.
    /// </summary>
    /// <param name="cell">The cell to add the error to.</param>
    /// <param name="errors">The dictionary to store the errors.</param>
    /// <param name="errorMessage">The error message to add.</param>
    /// <param name="isAddComment">Flag indicating whether to add a comment to the cell with the error message (optional, default is true).</param>
    public static async Task AddErrorAsync(Cell cell, Dictionary<string, string> errors, string errorMessage, bool isAddComment = true, CancellationToken cancellationToken = default)
    {
        StyleErrorCell(cell);
        errors.Add(cell.Name, errorMessage);

        //thêm message
        if (isAddComment)
        {
            await AddCommentAsync(cell, errorMessage, cancellationToken);
        }
    }

    /// <summary>
    /// Adds an error to the specified cell and stores the error message in the errors dictionary.
    /// </summary>
    /// <param name="cell">The cell to add the error to.</param>
    /// <param name="errorMessage">The error message to add.</param>
    /// <param name="isAddComment">Flag indicating whether to add a comment to the cell with the error message (optional, default is true).</param>
    public static async Task AddErrorAsync(Cell cell, string errorMessage, bool isAddComment = true, CancellationToken cancellationToken = default)
    {
        StyleErrorCell(cell);
        //thêm message
        if (isAddComment)
        {
            await AddCommentAsync(cell, errorMessage, cancellationToken);
        }
    }

    /// <summary>
    /// Styles the error cell by changing the background color to red.
    /// </summary>
    /// <param name="cell">The cell to style.</param>
    public static void StyleErrorCell(Cell cell)
    {
        // thay đổi màu nền
        var style = cell.GetStyle();
        style.ForegroundColor = System.Drawing.Color.Red;
        style.Pattern = BackgroundType.Solid;
        cell.SetStyle(style);
    }

    /// <summary>
    /// Adds a comment to the specified cell with the provided error message.
    /// </summary>
    /// <param name="cell">The cell to add the comment to.</param>
    /// <param name="errorMessage">The error message to include in the comment.</param>
    public static async Task AddCommentAsync(Cell cell, string errorMessage, CancellationToken cancellationToken = default)
    {
        await Task.Run(() =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var comment = cell.Worksheet.Comments[cell.Worksheet.Comments.Add(cell.Name)];
            comment.Note = errorMessage;

            // Adjust the size of the comment box based on the length of the error message
            int charWidth = 7; // Approximate width of a character in pixels
            int charHeight = 20; // Approximate height of a character in pixels
            int padding = 10; // Padding in pixels

            int lines = errorMessage.Split('\n').Length;
            int maxLineLength = errorMessage.Split('\n').Max(line => line.Length);

            comment.Width = (maxLineLength * charWidth) + padding;
            comment.Height = (lines * charHeight) + padding;
        }, cancellationToken);
    }

    /// <summary>
    /// Sets the value of a cell to the corresponding property of an instance.
    /// If the value cannot be converted, an error message is returned.
    /// </summary>
    /// <param name="value">The value of the cell.</param>
    /// <param name="instance">The instance of the object.</param>
    /// <param name="property">The property to set the value to.</param>
    /// <param name="uniqueValueTracker">The dictionary to track unique values.</param>
    /// <returns>An error message if the value cannot be converted, otherwise an empty string.</returns>
    private static async Task SetValueCellToPropertyAsync<T>(Cell cell, string sheetName, T instance, PropertyInfo property, Dictionary<string, HashSet<object>> uniqueValueTracker, bool isValidateData, bool isHeaderRequired, CancellationToken cancellationToken = default) where T : DocumentItemBase
    {
        await Task.Run(() =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            try
            {
                var value = cell.StringValue;
                object? currentValue = null;
                // lấy giá trị từ ô hiện tại và set vào property, nếu không convert được sẽ add error
                if (property.PropertyType == typeof(int))
                {
                    if (int.TryParse(value, out int intValue))
                    {
                        currentValue = intValue;
                        property.SetValue(instance, intValue);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeInteger);
                    }
                }
                else if (property.PropertyType == typeof(long))
                {
                    if (long.TryParse(value, out long intValue))
                    {
                        currentValue = intValue;
                        property.SetValue(instance, intValue);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeLong);
                    }
                }
                else if (property.PropertyType == typeof(float))
                {
                    if (float.TryParse(value, out float floatValue))
                    {
                        currentValue = floatValue;
                        property.SetValue(instance, floatValue);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeFloat);
                    }
                }
                else if (property.PropertyType == typeof(decimal))
                {
                    if (decimal.TryParse(value, out decimal decimalValue))
                    {
                        currentValue = decimalValue;
                        property.SetValue(instance, decimalValue);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeDecimal);
                    }
                }
                else if (property.PropertyType == typeof(DateTime))
                {
                    if (property.GetCustomAttributes(typeof(DateTimeAttribute), true).Any() && DateTime.TryParseExact(value, property.GetDateTimeFormat(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTimeConvertCheck))
                    {
                        currentValue = dateTimeConvertCheck;
                        property.SetValue(instance, dateTimeConvertCheck);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        if (DateTime.TryParseExact(value, CommonConstants.DefaultDateTimeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTimeConverted))
                        {
                            currentValue = dateTimeConverted;
                            property.SetValue(instance, dateTimeConverted);
                        }
                        else
                        {
                            instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeDateTime);
                        }
                    }
                }
                else if (property.PropertyType == typeof(DateOnly))
                {
                    if (property.GetCustomAttributes(typeof(DateAttribute), true).Any() && DateOnly.TryParseExact(value, property.GetDateFormat(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dateConvert))
                    {
                        currentValue = dateConvert;
                        property.SetValue(instance, dateConvert);
                    }
                    else if (!isValidateData)
                    {
                        property.SetValue(instance, DBNull.Value);
                    }
                    else
                    {
                        if (DateOnly.TryParseExact(value, CommonConstants.DefaultDateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dateConverted))
                        {
                            currentValue = dateConverted;
                            property.SetValue(instance, dateConverted);
                        }
                        else
                        {
                            instance.AddError(property.Name, sheetName, cell.Name, ErrorMessageConstants.InvalidTypeDate);
                        }
                    }
                }
                else
                {
                    currentValue = value;
                    property.SetValue(instance, value);
                }

                if (isValidateData)
                {
                    // Kiểm tra trùng lặp nếu thuộc tính có UniqueValuesAttribute
                    if (property.GetCustomAttributes(typeof(UniqueValuesAttribute), true).Any())
                    {
                        var propertyName = property.GetPropertyNameOrAlias().ToUpper();
                        if (!uniqueValueTracker.ContainsKey(propertyName))
                        {
                            uniqueValueTracker[propertyName] = new HashSet<object>();
                        }

                        if (!uniqueValueTracker[propertyName].Add(currentValue))
                        {
                            instance.AddError(property.Name, sheetName, cell.Name, string.Format(ErrorMessageConstants.DuplicateValueFound, value));
                        }
                    }
                    var test = property.GetValue(instance);
                    var test1 = property.GetValue(instance)?.ToString();
                    // kiểm tra required nếu có header required
                    if (isHeaderRequired && (property.GetValue(instance) == null || string.IsNullOrEmpty(property.GetValue(instance)?.ToString())))
                    {
                        instance.AddError(property.Name, sheetName, cell.Name, string.Format(ErrorMessageConstants.Required, value));
                    }
                    // validate property
                    var validationContext = new ValidationContext(instance)
                    {
                        MemberName = property.Name
                    };

                    // List to hold validation results
                    var validationResults = new List<ValidationResult>();

                    // Validate the property
                    bool isValid = Validator.TryValidateProperty(property.GetValue(instance), validationContext, validationResults);
                    if (!isValid)
                    {
                        foreach (var validationResult in validationResults)
                        {
                            foreach (var memberName in validationResult.MemberNames)
                            {
                                instance.AddError(memberName, sheetName, cell.Name, validationResult.ErrorMessage ?? string.Empty);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                instance.AddError(property.Name, sheetName, cell.Name, $"Error parsing value: {ex.Message}");
            }
        }, cancellationToken);
    }

    /// <summary>
    /// Adds a protected sheet to the workbook with the specified sheet name, password, key, and cell.
    /// </summary>
    /// <param name="data">The stream containing the workbook data.</param>
    /// <param name="sheetName">The name of the sheet to add.</param>
    /// <param name="password">The password to protect the sheet.</param>
    /// <param name="key">The key to validate.</param>
    /// <param name="cell">The cell containing the value to compare with the key.</param>
    /// <returns>The stream containing the updated workbook data.</returns>
    public static async Task<Stream> AddProtectedSheetAsync(Stream data, string sheetName, string password, string key, string cell, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            var workbook = new Workbook(data);

            var worksheetKey = workbook.Worksheets[sheetName];
            if (worksheetKey is not null)
            {
                data.Position = 0;
                return data;
            }

            // Thêm một sheet mới
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Đặt giá trị cho ô
            var cellKey = worksheet.Cells[cell];
            cellKey.PutValue(key);

            // Đặt màu nền cho ô
            var style = cellKey.GetStyle();
            style.Font.Color = System.Drawing.Color.White;
            style.ForegroundColor = System.Drawing.Color.White;
            style.Pattern = BackgroundType.Solid;
            cellKey.SetStyle(style);

            // Bảo vệ sheet
            // không cho xóa
            worksheet.Protection.AllowDeletingColumn = false;
            worksheet.Protection.AllowDeletingRow = false;

            // không cho sửa
            worksheet.Protection.AllowEditingContent = false;
            worksheet.Protection.AllowEditingObject = false;
            worksheet.Protection.AllowEditingScenario = false;

            //không cho thêm
            worksheet.Protection.AllowInsertingRow = false;

            // không cho lọc
            worksheet.Protection.AllowFiltering = false;

            //không cho sắp xếp
            worksheet.Protection.AllowSorting = false;

            // không cho đổi format
            worksheet.Protection.AllowFormattingCell = false;
            worksheet.Protection.AllowFormattingRow = false;
            worksheet.Protection.AllowFormattingColumn = false;

            // không cho thêm link
            worksheet.Protection.AllowInsertingHyperlink = false;

            // không cho select ô khóa
            worksheet.Protection.AllowSelectingLockedCell = false;

            // không cho select ô không khóa
            worksheet.Protection.AllowSelectingUnlockedCell = true;

            // không cho dùng pivot tables
            worksheet.Protection.AllowUsingPivotTable = false;

            //đặt mật khẩu
            worksheet.Protection.Password = password;

            // Ẩn sheet
            worksheet.IsVisible = false;

            //protect workbook - không cho thêm sửa xóa tất cả các sheet
            workbook.Protect(ProtectionType.Structure, password);
            // Lưu workbook vào stream
            var outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);
            outputStream.Position = 0;

            return outputStream;
        }, cancellationToken);

    }

    #endregion import

    #region export

    public static void GroupCell(Workbook wb, string rangeName, int[] columnGroup, GroupColumnItemTypeEnum groupColumnItemTypeEnum)
    {
        var worksheet = wb.Worksheets[0];
        var range1 = wb.Worksheets.GetRangeByName(rangeName);
        if (range1 is null || range1.RowCount <= 2)
        {
            return;
        }

        var cellGroups = columnGroup;
        var groups = new Dictionary<int, int>();
        if (groupColumnItemTypeEnum == GroupColumnItemTypeEnum.All)
        {
            foreach (var cellGroup in cellGroups)
            {
                worksheet.Cells.Merge(range1.FirstRow + 1, cellGroup, range1.RowCount - 1, 1);
            }

            return;
        }

        for (var j = range1.FirstColumn; j < range1.FirstColumn + 1; j++)
        {
            var firstIndex = 0;
            var lastIndex = 0;
            var cellFirst = worksheet.Cells[range1.FirstRow, j];
            var headerRowCount = !cellFirst.IsMerged ? 1 : cellFirst.GetMergedRange().RowCount;
            for (var i = range1.FirstRow + headerRowCount;
                 i < range1.FirstRow + (headerRowCount) + range1.RowCount;
                 i++)
            {
                var cell = worksheet.Cells[i, j];
                var cellValue = cell.Value?.ToString();
                var cellNext = worksheet.Cells[i + 1, j];
                var cellNextValue = cellNext.Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    firstIndex = i;
                }

                if (!string.IsNullOrEmpty(cellNextValue))
                {
                    lastIndex = i;
                    groups.TryAdd(firstIndex, lastIndex);

                    continue;
                }

                if (i != range1.FirstRow + range1.RowCount - 1)
                {
                    continue;
                }

                lastIndex = i;
                groups.TryAdd(firstIndex, lastIndex);
            }
        }

        foreach (var cellGroup in cellGroups)
        {
            foreach (var group in groups)
            {
                worksheet.Cells.Merge(group.Key, cellGroup, group.Value - group.Key + 1, 1);
            }
        }
    }

    public static async Task SetupImages(Dictionary<string, ImageDetail> images, WorkbookDesigner wd, HttpClient client)
    {
        var t = new DataTable("Images");
        foreach (var item in images)
        {
            // Add a column to save pictures.
            var dc = t.Columns.Add(item.Key);
            // Set its data type.
            dc.DataType = typeof(object);
        }

        var row = t.NewRow();
        foreach (var item in images)
        {
            if (string.IsNullOrEmpty(item.Value.Url))
            {
                continue;
            }

            try
            {
                var uri = new Uri(item.Value.Url);
                var bytes = await client.DownloadFileTaskAsync(uri);
                row[item.Key] = bytes;
            }
            catch
            {
                // ignored
            }
        }

        t.Rows.Add(row);
        wd.SetDataSource(t);
    }

    public static void SetResource(WorkbookDesigner designer, Dictionary<string, string> data, bool useCustomData)
    {
        foreach (var item in data)
        {
            var key = item.Key;
            var value = item.Value;
            var dataTable = JsonConvert.DeserializeObject<List<object>>(value);
            if (dataTable is null)
            {
                continue;
            }

            if (!useCustomData)
            {
                designer.SetDataSource(key, dataTable);
                continue;
            }

            var dataTable2 = new List<object>();
            foreach (JObject row in dataTable)
            {
                var maxItems = GetMaxItems(row);
                MapObjectItem(row, maxItems);
                var items = SplitItems(row, maxItems);
                dataTable2.AddRange(items);
            }

            designer.SetDataSource(key, dataTable2);
        }
    }

    public static void MapObjectItem(JObject jObject, int count)
    {
        foreach (var property in jObject.Properties())
        {
            var type = property.Value.Type;
            if (type != JTokenType.Array)
            {
                continue;
            }

            var items = ((JArray)property.Value);
            var currentCount = items.Count;
            if (currentCount == 0 || currentCount == count)
            {
                continue;
            }

            var firstItem = (JObject)items.First();
            var clone = (JObject)firstItem.DeepClone();

            var remainingCount = count - currentCount;
            for (var i = 0; i < remainingCount; i++)
            {
                var obj = new JObject();
                foreach (var p in clone.Properties())
                {
                    obj.Add(p.Name, string.Empty);
                }

                items.Add(obj);
            }
        }
    }

    public static List<object> SplitItems(JObject jObject, int count)
    {
        if (count == 0)
        {
            return [jObject];
        }

        var items = new List<object>();
        for (var i = 0; i < count; i++)
        {
            var obj = (JObject)jObject.DeepClone();
            RemoveItem(obj, i);
            items.Add(obj);
        }

        return items;
    }

    public static void RemoveItem(JObject jObject, int index)
    {
        foreach (var property in jObject.Properties())
        {
            var type = property.Value.Type;
            if (type != JTokenType.Array)
            {
                if (index != 0)
                {
                    property.Value = null;
                }

                continue;
            }

            var items = ((JArray)property.Value);
            var item1 = items.Skip(index).Take(1).Select(p => p.DeepClone()).FirstOrDefault();
            items.RemoveAll();
            if (item1 is not null)
            {
                items.Add(item1);
            }
        }
    }

    public static int GetMaxItems(JObject jObject)
    {
        return (from property in jObject.Properties()
                let type = property.Value.Type
                where type == JTokenType.Array
                select ((JArray)property.Value).Count).Prepend(0).Max();
    }

    #endregion export    

    #region common

    public static string GetCellName(int rowIndex, int colIndex)
    {
        return CellsHelper.CellIndexToName(rowIndex, colIndex);
    }

    #endregion common  
}

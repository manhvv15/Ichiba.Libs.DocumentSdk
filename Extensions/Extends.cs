using System.Reflection;
using Ichiba.Libs.DocumentSdk.Extend;

namespace Ichiba.Libs.DocumentSdk.Extensions;

public static class Extends
{
    public static async Task<byte[]> DownloadFileTaskAsync(this HttpClient client, Uri uri)
    {
        using Stream s = await client.GetStreamAsync(uri);
        using MemoryStream fs = new MemoryStream();
        await s.CopyToAsync(fs);
        fs.Position = 0L;
        return fs.ToByteArray();
    }

    public static byte[] ToByteArray(this Stream stream)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        if (stream is MemoryStream memoryStream)
        {
            if (memoryStream.TryGetBuffer(out var buffer))
            {
                return buffer.Array;
            }

            return memoryStream.ToArray();
        }

        using MemoryStream memoryStream2 = new MemoryStream();
        stream.CopyTo(memoryStream2);
        return memoryStream2.GetBuffer();
    }

    public static string GetPropertyNameOrAlias(this PropertyInfo property)
    {
        var nameAttribute = property.GetCustomAttribute<AliasAttribute>();
        return nameAttribute?.GetAlias() ?? property.Name;
    }
    public static string GetDateFormat(this PropertyInfo property)
    {
        var dateAttribute = property.GetCustomAttribute<DateAttribute>();
        return dateAttribute?.GetDateFormat();
    }

    public static string GetDateTimeFormat(this PropertyInfo property)
    {
        var dateTimeAttribute = property.GetCustomAttribute<DateTimeAttribute>();
        return dateTimeAttribute?.GetDateTimeFormat();
    }
}

using System.Reflection;

namespace QuickXL
{
    internal sealed class XLSheetHelper<TDto> where TDto : class, new()
    {
        internal static Dictionary<string, string> GetHeaders()
        {
            return typeof(TDto).GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(XLHeader)))
                .ToDictionary(
                    prop => prop.Name,
                    prop => prop.GetCustomAttribute<XLHeader>()?.HeaderName ?? prop.Name);
        }

        internal static void MapPocoToSheet(TDto item, int rowIndex, Dictionary<string, string> headers, IExcelSheet sheet)
        {
            int columnIndex = 0;
            foreach (var header in headers)
            {
                var propertyValue = item.GetType().GetProperty(header.Key)?.GetValue(item)?.ToString() ?? string.Empty;
                sheet.AddCell(rowIndex, columnIndex++, propertyValue);
            }
        }
    }
}

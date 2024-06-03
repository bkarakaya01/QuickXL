using System.Reflection;

namespace QuickXL
{
    internal sealed class XLSheetHelper<TPoco> where TPoco : class, IExcelPOCO, new()
    {
        internal static Dictionary<string, string> GetHeaders()
        {
            return typeof(TPoco).GetProperties()
                .Where(prop => prop.GetCustomAttributes(typeof(ExportHeader), true).Length != 0)
                .ToDictionary(
                    prop => prop.Name,
                    prop => prop.GetCustomAttribute<ExportHeader>()?.HeaderName ?? prop.Name);
        }

        internal static void MapPocoToSheet(TPoco item, int rowIndex, Dictionary<string, string> headers, IExcelSheet sheet)
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

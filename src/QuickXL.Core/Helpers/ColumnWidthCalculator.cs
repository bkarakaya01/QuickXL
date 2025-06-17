using Ardalis.GuardClauses;

namespace QuickXL.Core.Helpers;

/// <summary>
/// Computes approximate column widths based on header and data lengths.
/// </summary>
internal static class ColumnWidthCalculator
{
    /// <summary>
    /// Estimates width for an Excel column.
    /// </summary>
    public static double Calculate<TDto>(
        IEnumerable<TDto> data,
        Func<TDto, object?> getter,
        int headerLength)
    {
        Guard.Against.Null(getter, nameof(getter));
        int max = headerLength;
        foreach (var item in data)
        {
            var text = getter(item)?.ToString() ?? string.Empty;
            if (text.Length > max) max = text.Length;
        }
        return Math.Min(max * 1.2, 255);
    }
}

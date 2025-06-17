using Ardalis.GuardClauses;

namespace QuickXL.Core.Helpers;

/// <summary>
/// Provides Excel column name and cell reference utilities.
/// </summary>
internal static class ExcelReferenceHelper
{
    /// <summary>
    /// Converts zero-based column index to Excel column name (A, B, ..., Z, AA, ...).
    /// </summary>
    public static string GetColumnName(int index)
    {
        Guard.Against.Negative(index, nameof(index));
        string name = string.Empty;
        int dividend = index + 1;
        while (dividend > 0)
        {
            int mod = (dividend - 1) % 26;
            name = (char)('A' + mod) + name;
            dividend = (dividend - mod) / 26;
        }
        return name;
    }

    /// <summary>
    /// Builds a full cell reference (e.g. A1) from column name and one-based row index.
    /// </summary>
    public static string GetCellReference(string columnName, int row)
    {
        Guard.Against.NullOrEmpty(columnName, nameof(columnName));
        Guard.Against.Negative(row, nameof(row));
        return $"{columnName}{row}";
    }
}

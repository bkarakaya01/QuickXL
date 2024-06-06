using NPOI.SS.UserModel;
using QuickXL.Core.Contracts.Models;

namespace QuickXL.Core.Models.Cells;

// <summary>
/// Object as a reference to a cell in an excel file.
/// </summary>
internal class XLCell : IExcelUnit
{
    /// <summary>
    /// Cell value.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// High level representation of a cell.
    /// </summary>
    public ICell? Cell { get; set; }    
}

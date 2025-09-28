namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a worksheet within a workbook. Contains rows of data.
/// </summary>
public abstract class Worksheet
{
    /// <summary>
    /// Gets the name of the worksheet.
    /// </summary>
    public abstract string Name { get; }

    /// <summary>
    /// Gets an asynchronous enumerable collection of all rows in the worksheet.
    /// </summary>
    public abstract RowCollection Rows { get; }

    /// <summary>
    /// Synchronously gets a cell by its row and column numeric indices.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the cell's row.</param>
    /// <param name="columnIndex">The zero-based index of the cell's column.</param>
    /// <returns>The Cell at the specified row and column position.</returns>
    public Cell GetCell(int rowIndex, int columnIndex) => Rows[rowIndex].Cells[columnIndex];

    /// <summary>
    /// Synchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The Cell corresponding to the specified reference.</returns>
    public Cell GetCell(ReadOnlySpan<char> reference)
    {
        var (rowIndex, columnIndex) = SpreadsheetUtils.CellIndex(reference);
        return GetCell(rowIndex, columnIndex);
    }

    /// <summary>
    /// Asynchronously gets a cell by its row and column indices.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the cell's row.</param>
    /// <param name="columnIndex">The zero-based index of the cell's column.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Cell&gt; representing the asynchronous operation, containing the requested Cell.</returns>
    public async Task<Cell> GetCellAsync(int rowIndex, int columnIndex, CancellationToken cancellationToken = default)
    {
        var row = await Rows.GetItemAsync(rowIndex, cancellationToken);
        return row.Cells[columnIndex];
    }

    /// <summary>
    /// Asynchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Cell&gt; representing the asynchronous operation, containing the requested Cell.</returns>
    public Task<Cell> GetCellAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default)
    {
        var (rowIndex, columnIndex) = SpreadsheetUtils.CellIndex(reference);
        return GetCellAsync(rowIndex, columnIndex, cancellationToken);
    }
}

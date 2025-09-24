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
    public abstract IAsyncEnumerable<Row> Rows { get; }

    /// <summary>
    /// Synchronously gets a row by its numeric index.
    /// </summary>
    /// <param name="index">The zero-based index of the row to retrieve.</param>
    /// <returns>The Row at the specified index.</returns>
    public abstract Row GetRow(int index);

    /// <summary>
    /// Synchronously gets a row by its reference identifier.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <returns>The Row corresponding to the specified reference.</returns>
    public abstract Row GetRow(ReadOnlySpan<char> reference);

    /// <summary>
    /// Asynchronously gets a row by its index.
    /// </summary>
    /// <param name="index">The zero-based index of the row to retrieve.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Row&gt; representing the asynchronous operation, containing the requested Row.</returns>
    public abstract Task<Row> GetRowAsync(int index, CancellationToken cancellationToken = default);

    /// <summary>
    /// Asynchronously gets a row by its reference identifier.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Row&gt; representing the asynchronous operation, containing the requested Row.</returns>
    public abstract Task<Row> GetRowAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default);

    /// <summary>
    /// Synchronously gets a cell by its row and column numeric indices.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the cell's row.</param>
    /// <param name="columnIndex">The zero-based index of the cell's column.</param>
    /// <returns>The Cell at the specified row and column position.</returns>
    public abstract Cell GetCell(int rowIndex, int columnIndex);

    /// <summary>
    /// Synchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The Cell corresponding to the specified reference.</returns>
    public abstract Cell GetCell(ReadOnlySpan<char> reference);

    /// <summary>
    /// Asynchronously gets a cell by its row and column indices.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the cell's row.</param>
    /// <param name="columnIndex">The zero-based index of the cell's column.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Cell&gt; representing the asynchronous operation, containing the requested Cell.</returns>
    public abstract Task<Cell> GetCellAsync(int rowIndex, int columnIndex, CancellationToken cancellationToken = default);

    /// <summary>
    /// Asynchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Cell&gt; representing the asynchronous operation, containing the requested Cell.</returns>
    public abstract Task<Cell> GetCellAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default);
}

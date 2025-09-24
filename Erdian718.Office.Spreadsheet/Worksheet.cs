namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a worksheet within a workbook. Contains rows of data.
/// </summary>
public abstract class Worksheet
{
    private int _rowIndex = -1;
    private IAsyncEnumerator<Row>? _rowEnumerator;

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
    public virtual Row GetRow(int index) =>
        GetRowAsync(index).ConfigureAwait(false).GetAwaiter().GetResult();

    /// <summary>
    /// Synchronously gets a row by its reference identifier.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <returns>The Row corresponding to the specified reference.</returns>
    public virtual Row GetRow(ReadOnlySpan<char> reference) =>
        GetRow(SpreadsheetUtils.RowIndex(reference));

    /// <summary>
    /// Asynchronously gets a row by its index.
    /// </summary>
    /// <param name="index">The zero-based index of the row to retrieve.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Row&gt; representing the asynchronous operation, containing the requested Row.</returns>
    public virtual async Task<Row> GetRowAsync(int index, CancellationToken cancellationToken = default)
    {
        if (index < 0) throw new IndexOutOfRangeException($"Invalid row index [{index}].");

        if (index < _rowIndex)
        {
            await _rowEnumerator!.DisposeAsync().ConfigureAwait(false);
            _rowEnumerator = null;
            _rowIndex = -1;
        }

        var enumerator = _rowEnumerator;
        if (enumerator is null)
        {
            enumerator = Rows.GetAsyncEnumerator(cancellationToken);
            _rowEnumerator = enumerator;
        }

        while (await enumerator.MoveNextAsync())
        {
            _rowIndex++;
            if (index == _rowIndex)
            {
                return enumerator.Current;
            }
        }
        throw new IndexOutOfRangeException($"Row index [{index}] is out of range.");
    }

    /// <summary>
    /// Asynchronously gets a row by its reference identifier.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Row&gt; representing the asynchronous operation, containing the requested Row.</returns>
    public virtual Task<Row> GetRowAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default) =>
        GetRowAsync(SpreadsheetUtils.RowIndex(reference), cancellationToken);

    /// <summary>
    /// Synchronously gets a cell by its row and column numeric indices.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the cell's row.</param>
    /// <param name="columnIndex">The zero-based index of the cell's column.</param>
    /// <returns>The Cell at the specified row and column position.</returns>
    public virtual Cell GetCell(int rowIndex, int columnIndex) =>
        GetRow(rowIndex).Cells[columnIndex];

    /// <summary>
    /// Synchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The Cell corresponding to the specified reference.</returns>
    public virtual Cell GetCell(ReadOnlySpan<char> reference)
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
    public virtual async Task<Cell> GetCellAsync(int rowIndex, int columnIndex, CancellationToken cancellationToken = default)
    {
        var row = await GetRowAsync(rowIndex, cancellationToken).ConfigureAwait(false);
        return row.Cells[columnIndex];
    }

    /// <summary>
    /// Asynchronously gets a cell by its reference identifier.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>A Task&lt;Cell&gt; representing the asynchronous operation, containing the requested Cell.</returns>
    public virtual Task<Cell> GetCellAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default)
    {
        var (rowIndex, columnIndex) = SpreadsheetUtils.CellIndex(reference);
        return GetCellAsync(rowIndex, columnIndex, cancellationToken);
    }
}

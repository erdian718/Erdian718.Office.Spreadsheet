namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of rows in a worksheet.
/// </summary>
/// <param name="rows">The rows in the collection.</param>
public class RowCollection(IAsyncEnumerable<Row> rows) : IAsyncEnumerable<Row>
{
    private static readonly Row s_emptyRow = new EmptyRow();

    private int _index = -1;
    private IAsyncEnumerator<Row>? _enumerator;
    private readonly IAsyncEnumerable<Row> _rows = rows;

    /// <summary>
    /// Gets a row from the collection by row index.
    /// </summary>
    /// <param name="index">The zero-based index of the row to retrieve.</param>
    /// <returns>The Row at the specified index.</returns>
    public Row this[int index] => GetItemAsync(index).ConfigureAwait(false).GetAwaiter().GetResult();

    /// <summary>
    /// Gets a row from the collection by row reference.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <returns>The Row corresponding to the specified reference.</returns>
    public Row this[ReadOnlySpan<char> reference] => this[SpreadsheetUtils.RowIndex(reference)];

    /// <summary>
    /// Asynchronously gets a row from the collection by row index.
    /// </summary>
    /// <param name="reference">The zero-based index of the row to retrieve.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A Task&lt;Row&gt; that represents the asynchronous operation.</returns>
    public async Task<Row> GetItemAsync(int index, CancellationToken cancellationToken = default)
    {
        var enumerator = GetAsyncEnumeratorAt(index, cancellationToken);
        if (await enumerator.MoveNextAsync())
        {
            return enumerator.Current;
        }
        return s_emptyRow;
    }

    /// <summary>
    /// Asynchronously gets a row from the collection by row reference.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A Task&lt;Row&gt; that represents the asynchronous operation.</returns>
    public Task<Row> GetItemAsync(ReadOnlySpan<char> reference, CancellationToken cancellationToken = default) =>
        GetItemAsync(SpreadsheetUtils.RowIndex(reference), cancellationToken);

    /// <summary>
    /// Returns an asynchronous enumerator that iterates through the collection of rows.
    /// </summary>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>An IAsyncEnumerator&lt;Row&gt; that can be used to iterate through the collection.</returns>
    public IAsyncEnumerator<Row> GetAsyncEnumerator(CancellationToken cancellationToken = default) =>
        _rows.GetAsyncEnumerator(cancellationToken);

    /// <summary>
    /// Asynchronously gets an asynchronous enumerator that iterates over Row items starting from the specified index.
    /// </summary>
    /// <param name="index">The zero-based starting index from which to begin enumerating rows.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous enumeration operation.</param>
    /// <returns>An <see cref="IAsyncEnumerator{Row}"/> that can be used to asynchronously iterate over rows starting from the specified index.</returns>
    /// <remarks>Callers should not dispose of the returned <see cref="IAsyncEnumerator{Row}"/>, as they are managed internally by the containing class.</remarks>
    internal async IAsyncEnumerator<Row> GetAsyncEnumeratorAt(int index, CancellationToken cancellationToken = default)
    {
        if (index < 0) throw new ArgumentOutOfRangeException($"Invalid row index [{index}].");
        if (index < _index)
        {
            await _enumerator!.DisposeAsync().ConfigureAwait(false);
            _index = -1;
            _enumerator = null;
        }

        var enumerator = _enumerator;
        if (enumerator is null)
        {
            enumerator = _rows.GetAsyncEnumerator(cancellationToken);
            _enumerator = enumerator;
        }
        if (index == _index) yield return enumerator.Current;

        index -= 1;
        while (index > _index)
        {
            if (!await enumerator.MoveNextAsync()) break;
            _index++;
        }

        while (await enumerator.MoveNextAsync())
        {
            _index++;
            yield return enumerator.Current;
        }
    }

    /// <summary>
    /// Asynchronously disposes the resources used by the RowCollection.
    /// </summary>
    /// <returns>A ValueTask that represents the asynchronous dispose operation.</returns>
    internal async ValueTask DisposeAsync()
    {
        var enumerator = _enumerator;
        if (enumerator is not null)
        {
            await enumerator.DisposeAsync().ConfigureAwait(false);
            _index = -1;
            _enumerator = null;
        }
    }

    /// <summary>
    /// Represents an empty row.
    /// </summary>
    internal class EmptyRow : Row
    {
        private static readonly CellCollection s_cells = new([]);

        /// <summary>
        /// Gets a value indicating whether the row is blank.
        /// </summary>
        public override bool IsHidden => false;

        /// <summary>
        /// Gets a collection of cells in the row.
        /// </summary>
        public override CellCollection Cells => s_cells;
    }
}

using System.Collections;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of rows in a worksheet.
/// </summary>
public abstract class RowCollection : IEnumerable<Row>, IAsyncEnumerable<Row>
{
    /// <summary>
    /// Gets a row from the collection by row index.
    /// </summary>
    /// <param name="index">The zero-based index of the row to retrieve.</param>
    /// <returns>The Row at the specified index.</returns>
    public Row this[int index]
    {
        get
        {
            var enumerator = GetEnumeratorAt(index);
            return enumerator.MoveNext() ? enumerator.Current : Row.Empty;
        }
    }

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
        return await enumerator.MoveNextAsync() ? enumerator.Current : Row.Empty;
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
    /// Returns an enumerator that iterates through the collection of rows.
    /// </summary>
    /// <returns>An IEnumerator&lt;Row&gt; that can be used to iterate through the collection.</returns>
    public abstract IEnumerator<Row> GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the collection of rows.
    /// </summary>
    /// <returns>An IEnumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Returns an asynchronous enumerator that iterates through the collection of rows.
    /// </summary>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the operation.</param>
    /// <returns>An IAsyncEnumerator&lt;Row&gt; that can be used to iterate through the collection.</returns>
    public abstract IAsyncEnumerator<Row> GetAsyncEnumerator(CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets an enumerator that iterates over Row items starting from the specified index.
    /// </summary>
    /// <param name="index">The zero-based starting index from which to begin enumerating rows.</param>
    /// <returns>An <see cref="IEnumerator{Row}"/> that can be used to iterate over rows starting from the specified index.</returns>
    /// <remarks>Callers should not dispose of the returned <see cref="IEnumerator{Row}"/>, as they are managed internally by the containing class.</remarks>
    internal protected abstract IEnumerator<Row> GetEnumeratorAt(int index);

    /// <summary>
    /// Gets an asynchronous enumerator that iterates over Row items starting from the specified index.
    /// </summary>
    /// <param name="index">The zero-based starting index from which to begin enumerating rows.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous enumeration operation.</param>
    /// <returns>An <see cref="IAsyncEnumerator{Row}"/> that can be used to asynchronously iterate over rows starting from the specified index.</returns>
    /// <remarks>Callers should not dispose of the returned <see cref="IAsyncEnumerator{Row}"/>, as they are managed internally by the containing class.</remarks>
    internal protected abstract IAsyncEnumerator<Row> GetAsyncEnumeratorAt(int index, CancellationToken cancellationToken = default);
}

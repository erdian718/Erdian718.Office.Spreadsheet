using System.Collections;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of cells in a row.
/// </summary>
public abstract class CellCollection : IEnumerable<Cell>
{
    /// <summary>
    /// Gets a cell from the collection by column index.
    /// </summary>
    /// <param name="index">The zero-based index of the cell to retrieve.</param>
    /// <returns>The Cell at the specified index.</returns>
    public abstract Cell this[int index] { get; }

    /// <summary>
    /// Gets a cell from the collection by column reference.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The Cell corresponding to the specified reference.</returns>
    public virtual Cell this[ReadOnlySpan<char> reference] => this[SpreadsheetUtils.ColumnIndex(reference)];

    /// <summary>
    /// Returns an enumerator that iterates through the collection of cells.
    /// </summary>
    /// <returns>An IEnumerator&lt;Cell&gt; that can be used to iterate through the collection.</returns>
    public abstract IEnumerator<Cell> GetEnumerator();

    /// <summary>
    /// Explicit implementation of IEnumerable.GetEnumerator.
    /// </summary>
    /// <returns>An IEnumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

using System.Collections;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of cells in a row.
/// </summary>
/// <param name="cells">The cells in the collection.</param>
public class CellCollection(IEnumerable<Cell> cells) : IEnumerable<Cell>
{
    private static readonly Cell s_emptyCell = new EmptyCell();

    private readonly Cell[] _cells = [.. cells];

    /// <summary>
    /// Gets a cell from the collection by column index.
    /// </summary>
    /// <param name="index">The zero-based index of the cell to retrieve.</param>
    /// <returns>The Cell at the specified index.</returns>
    public Cell this[int index] => index < _cells.Length ? _cells[index] : s_emptyCell;

    /// <summary>
    /// Gets a cell from the collection by column reference.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The Cell corresponding to the specified reference.</returns>
    public Cell this[ReadOnlySpan<char> reference] => this[SpreadsheetUtils.ColumnIndex(reference)];

    /// <summary>
    /// Gets the number of items contained in the collection.
    /// </summary>
    public int Length => _cells.Length;

    /// <summary>
    /// Returns an enumerator that iterates through the collection of cells.
    /// </summary>
    /// <returns>An IEnumerator&lt;Cell&gt; that can be used to iterate through the collection.</returns>
    public IEnumerator<Cell> GetEnumerator() => ((IEnumerable<Cell>)_cells).GetEnumerator();

    /// <summary>
    /// Explicit implementation of IEnumerable.GetEnumerator.
    /// </summary>
    /// <returns>An IEnumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => _cells.GetEnumerator();

    /// <summary>
    /// Represents an empty cell.
    /// </summary>
    internal class EmptyCell : Cell
    {
        /// <summary>
        /// Gets the underlying raw value stored in the cell.
        /// </summary>
        public override object? Value => null;
    }
}

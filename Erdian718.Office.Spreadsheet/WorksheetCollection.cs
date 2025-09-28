using System.Collections;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of worksheets in a workbook.
/// </summary>
public abstract class WorksheetCollection : IEnumerable<Worksheet>
{
    /// <summary>
    /// Gets a worksheet from the collection by its numeric index.
    /// </summary>
    /// <param name="index">The zero-based index of the worksheet to retrieve.</param>
    /// <returns>The Worksheet at the specified index.</returns>
    public virtual Worksheet this[int index] =>
        this.ElementAtOrDefault(index) ?? throw new InvalidOperationException($"Worksheet [{index}] is not found.");

    /// <summary>
    /// Gets a worksheet from the collection by its name.
    /// </summary>
    /// <param name="name">The name of the worksheet to retrieve.</param>
    /// <returns>The Worksheet with the specified name.</returns>
    public virtual Worksheet this[ReadOnlySpan<char> name]
    {
        get
        {
            foreach (var worksheet in this)
            {
                if (name.SequenceEqual(worksheet.Name))
                {
                    return worksheet;
                }
            }
            throw new InvalidOperationException($"Worksheet [{name}] is not found.");
        }
    }

    /// <summary>
    /// Gets the number of items contained in the collection.
    /// </summary>
    public abstract int Length { get; }

    /// <summary>
    /// Returns an enumerator that iterates through the collection of worksheets.
    /// </summary>
    /// <returns>An IEnumerator&lt;Worksheet&gt; that can be used to iterate through the collection.</returns>
    public abstract IEnumerator<Worksheet> GetEnumerator();

    /// <summary>
    /// Explicit implementation of IEnumerable.GetEnumerator.
    /// </summary>
    /// <returns>An IEnumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

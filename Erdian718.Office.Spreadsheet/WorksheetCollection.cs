using System.Collections;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a collection of worksheets in a workbook.
/// </summary>
/// <param name="worksheets">The worksheets in the workbook.</param>
public class WorksheetCollection(IEnumerable<Worksheet> worksheets) : IEnumerable<Worksheet>
{
    private readonly Worksheet[] _worksheets = [.. worksheets];

    /// <summary>
    /// Gets a worksheet from the collection by its numeric index.
    /// </summary>
    /// <param name="index">The zero-based index of the worksheet to retrieve.</param>
    /// <returns>The Worksheet at the specified index.</returns>
    public Worksheet this[int index] => _worksheets[index];

    /// <summary>
    /// Gets a worksheet from the collection by its name.
    /// </summary>
    /// <param name="name">The name of the worksheet to retrieve.</param>
    /// <returns>The Worksheet with the specified name.</returns>
    public Worksheet this[ReadOnlySpan<char> name]
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
    public int Length => _worksheets.Length;

    /// <summary>
    /// Returns an enumerator that iterates through the collection of worksheets.
    /// </summary>
    /// <returns>An IEnumerator&lt;Worksheet&gt; that can be used to iterate through the collection.</returns>
    public IEnumerator<Worksheet> GetEnumerator() => ((IEnumerable<Worksheet>)_worksheets).GetEnumerator();

    /// <summary>
    /// Explicit implementation of IEnumerable.GetEnumerator.
    /// </summary>
    /// <returns>An IEnumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => _worksheets.GetEnumerator();
}

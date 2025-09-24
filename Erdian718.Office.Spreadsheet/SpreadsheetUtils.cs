using System.Globalization;

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Provides utility methods for spreadsheet.
/// </summary>
public static class SpreadsheetUtils
{
    /// <summary>
    /// Converts a row index to its corresponding reference.
    /// </summary>
    /// <param name="index">The zero-based row index.</param>
    /// <returns>A row reference.</returns>
    public static string RowReference(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException($"Invalid row index: {index}.");
        return (index + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Converts a column index to its corresponding reference.
    /// </summary>
    /// <param name="index">The zero-based column index.</param>
    /// <returns>A column reference.</returns>
    public static string ColumnReference(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException($"Invalid column index: {index}.");
        if (index < 26) return ((char)('A' + index)).ToString(CultureInfo.InvariantCulture);
        return ColumnReference(index / 26 - 1) + ColumnReference(index % 26);
    }

    /// <summary>
    /// Converts the cell indices to their corresponding reference.
    /// </summary>
    /// <param name="rowIndex">The zero-based row index.</param>
    /// <param name="columnIndex">The zero-based column index.</param>
    /// <returns>A cell reference.</returns>
    public static string CellReference(int rowIndex, int columnIndex) =>
        ColumnReference(columnIndex) + RowReference(rowIndex);

    /// <summary>
    /// Converts the range indices to their corresponding reference.
    /// </summary>
    /// <param name="rowIndex1">The zero-based row index of the first cell in the range.</param>
    /// <param name="columnIndex1">The zero-based column index of the first cell in the range.</param>
    /// <param name="rowIndex2">The zero-based row index of the second cell in the range.</param>
    /// <param name="columnIndex2">The zero-based column index of the second cell in the range.</param>
    /// <returns>A range reference.</returns>
    public static string RangeReference(int rowIndex1, int columnIndex1, int rowIndex2, int columnIndex2) =>
        CellReference(rowIndex1, columnIndex1) + ":" + CellReference(rowIndex2, columnIndex2);

    /// <summary>
    /// Converts a row reference to its corresponding index.
    /// </summary>
    /// <param name="reference">The row reference.</param>
    /// <returns>A zero-based row index.</returns>
    public static int RowIndex(ReadOnlySpan<char> reference)
    {
        if (!int.TryParse(reference, out var index) || index < 1)
        {
            throw new ArgumentException($"Invalid row reference: {reference}.");
        }
        return index - 1;
    }

    /// <summary>
    /// Converts a column reference to its corresponding index.
    /// </summary>
    /// <param name="reference">The column reference.</param>
    /// <returns>A zero-based column index.</returns>
    public static int ColumnIndex(ReadOnlySpan<char> reference)
    {
        var n = reference.Length;
        if (n < 1) throw new ArgumentException("Empty column reference.");

        if (n == 1)
        {
            var c = reference[0];
            if (c >= 'A' && c <= 'Z') return c - 'A';
            if (c >= 'a' && c <= 'z') return c - 'a';
            throw new ArgumentException($"Invalid column reference: {reference}.");
        }

        n -= 1;
        var a = ColumnIndex(reference[n..]);
        var b = ColumnIndex(reference[..n]);
        return a + 26 * b + 26;
    }

    /// <summary>
    /// Converts a cell reference to its corresponding indices.
    /// </summary>
    /// <param name="reference">The cell reference.</param>
    /// <returns>The zero-based cell indices.</returns>
    public static (int RowIndex, int ColumnIndex) CellIndex(ReadOnlySpan<char> reference)
    {
        for (var k = 0; k < reference.Length; k++)
        {
            var c = reference[k];
            if (c >= '0' && c <= '9') return (RowIndex(reference[k..]), ColumnIndex(reference[..k]));
        }
        throw new ArgumentException($"Invalid cell reference: {reference}.");
    }

    /// <summary>
    /// Converts a range reference to its corresponding indices.
    /// </summary>
    /// <param name="reference">The range reference.</param>
    /// <returns>The zero-based range indices.</returns>
    public static (int RowIndex1, int ColumnIndex1, int RowIndex2, int ColumnIndex2) RangeIndex(ReadOnlySpan<char> reference)
    {
        var k = reference.IndexOf(':');
        if (k < 0) throw new ArgumentException($"Invalid range reference: {reference}.");

        var (rowIndex1, columnIndex1) = CellIndex(reference[..k]);
        var (rowIndex2, columnIndex2) = CellIndex(reference[(k + 1)..]);
        if (rowIndex2 < rowIndex1 || columnIndex2 < columnIndex1) throw new ArgumentException($"Invalid range reference: {reference}.");
        return (rowIndex1, columnIndex1, rowIndex2, columnIndex2);
    }
}

namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a row in a worksheet. Contains a collection of cells, organized by column position.
/// </summary>
public abstract class Row
{
    /// <summary>
    /// Gets a value indicating whether the row is blank.
    /// </summary>
    public abstract bool IsBlank { get; }

    /// <summary>
    /// Gets a value indicating whether the row is hidden in the worksheet view.
    /// </summary>
    public abstract bool IsHidden { get; }
}

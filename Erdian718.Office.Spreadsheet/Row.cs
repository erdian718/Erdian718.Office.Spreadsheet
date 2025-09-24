namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a row in a worksheet. Contains a collection of cells, organized by column position.
/// </summary>
public abstract class Row
{
    /// <summary>
    /// Gets a value indicating whether the row is blank.
    /// </summary>
    public virtual bool IsBlank => Cells.All(cell => cell.IsBlank);

    /// <summary>
    /// Gets a value indicating whether the row is hidden in the worksheet view.
    /// </summary>
    public abstract bool IsHidden { get; }

    /// <summary>
    /// Gets a collection of cells in the row.
    /// </summary>
    public abstract CellCollection Cells { get; }
}

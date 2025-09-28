namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a row in a worksheet. Contains a collection of cells, organized by column position.
/// </summary>
public class Row
{
    /// <summary>
    /// Represents a empty row.
    /// </summary>
    internal protected static Row Empty { get; } = new();

    /// <summary>
    /// Gets a value indicating whether the row is blank.
    /// </summary>
    public bool IsBlank => Cells.All(cell => cell.IsBlank);

    /// <summary>
    /// Gets a value indicating whether the row is hidden in the worksheet view.
    /// </summary>
    public virtual bool IsHidden => false;

    /// <summary>
    /// Gets a collection of cells in the row.
    /// </summary>
    public virtual CellCollection Cells { get; } = new();
}

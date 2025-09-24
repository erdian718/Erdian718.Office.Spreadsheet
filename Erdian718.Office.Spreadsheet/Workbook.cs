namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a spreadsheet workbook. Acts as a top-level container for managing worksheets.
/// </summary>
public abstract class Workbook
{
    /// <summary>
    /// Gets the collection of worksheets contained in the current workbook.
    /// </summary>
    public abstract WorksheetCollection Worksheets { get; }

    /// <summary>
    /// Synchronously saves the workbook content to the specified stream.
    /// </summary>
    /// <param name="stream">The target stream to receive the workbook content.</param>
    public abstract void Save(Stream stream);

    /// <summary>
    /// Synchronously saves the workbook content to the specified file path.
    /// </summary>
    /// <param name="path">The target file path where the workbook will be saved.</param>
    public abstract void Save(string path);

    /// <summary>
    /// Asynchronously saves the workbook content to the specified stream.
    /// </summary>
    /// <param name="stream">The target stream to receive the workbook content.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the save operation.</param>
    /// <returns>A Task representing the asynchronous save operation.</returns>
    public abstract Task SaveAsync(Stream stream, CancellationToken cancellationToken = default);

    /// <summary>
    /// Asynchronously saves the workbook content to the specified file path.
    /// </summary>
    /// <param name="path">The target file path where the workbook will be saved.</param>
    /// <param name="cancellationToken">The Cancellation token that can be used to cancel the save operation.</param>
    /// <returns>A Task representing the asynchronous save operation.</returns>
    public abstract Task SaveAsync(string path, CancellationToken cancellationToken = default);
}

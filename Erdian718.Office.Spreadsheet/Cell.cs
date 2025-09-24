namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a cell in a row. Stores a single data value.
/// </summary>
public abstract class Cell
{
    /// <summary>
    /// Gets a value indicating whether the cell is blank.
    /// </summary>
    public abstract bool IsBlank { get; }

    /// <summary>
    /// Gets the underlying raw value stored in the cell.
    /// </summary>
    public abstract object? Value { get; }

    /// <summary>
    /// Retrieves the cell's value as a string.
    /// </summary>
    /// <returns>The cell's value as a string.</returns>
    public abstract string GetString();

    /// <summary>
    /// Retrieves the cell's value as a char.
    /// </summary>
    /// <returns>The cell's value as a char.</returns>
    public abstract char GetChar();

    /// <summary>
    /// Retrieves the cell's value as a boolean.
    /// </summary>
    /// <returns>The cell's value as a boolean.</returns>
    public abstract bool GetBoolean();

    /// <summary>
    /// Retrieves the cell's value as a single-precision floating-point number.
    /// </summary>
    /// <returns>The cell's value as a single-precision floating-point number.</returns>
    public abstract float GetSingle();

    /// <summary>
    /// Retrieves the cell's value as a double-precision floating-point number.
    /// </summary>
    /// <returns>The cell's value as a double-precision floating-point number.</returns>
    public abstract double GetDouble();

    /// <summary>
    /// Retrieves the cell's value as a decimal.
    /// </summary>
    /// <returns>The cell's value as a decimal.</returns>
    public abstract decimal GetDecimal();

    /// <summary>
    /// Retrieves the cell's value as a 8-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 8-bit signed integer.</returns>
    public abstract sbyte GetSByte();

    /// <summary>
    /// Retrieves the cell's value as a 8-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 8-bit unsigned integer.</returns>
    public abstract byte GetByte();

    /// <summary>
    /// Retrieves the cell's value as a 16-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 16-bit signed integer.</returns>
    public abstract short GetInt16();

    /// <summary>
    /// Retrieves the cell's value as a 16-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 16-bit unsigned integer.</returns>
    public abstract ushort GetUInt16();

    /// <summary>
    /// Retrieves the cell's value as a 32-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 32-bit signed integer.</returns>
    public abstract int GetInt32();

    /// <summary>
    /// Retrieves the cell's value as a 32-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 32-bit unsigned integer.</returns>
    public abstract uint GetUInt32();

    /// <summary>
    /// Retrieves the cell's value as a 64-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 64-bit signed integer.</returns>
    public abstract long GetInt64();

    /// <summary>
    /// Retrieves the cell's value as a 64-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 64-bit unsigned integer.</returns>
    public abstract ulong GetUInt64();

    /// <summary>
    /// Retrieves the cell's value as a DateTime object.
    /// </summary>
    /// <returns>The cell's value as a DateTime object.</returns>
    public abstract DateTime GetDateTime();

    /// <summary>
    /// Retrieves the cell's value as a DateOnly object.
    /// </summary>
    /// <returns>The cell's value as a DateOnly object.</returns>
    public abstract DateOnly GetDateOnly();

    /// <summary>
    /// Retrieves the cell's value as a TimeOnly object.
    /// </summary>
    /// <returns>The cell's value as a TimeOnly object.</returns>
    public abstract TimeOnly GetTimeOnly();

    /// <summary>
    /// Retrieves the cell's value as a DateTimeOffset object.
    /// </summary>
    /// <returns>The cell's value as a DateTimeOffset object.</returns>
    public abstract DateTimeOffset GetDateTimeOffset();

    /// <summary>
    /// Retrieves the cell's value as a string, or null.
    /// </summary>
    /// <returns>The cell's value as a string, or null.</returns>
    public abstract string? GetStringOrNull();

    /// <summary>
    /// Retrieves the cell's value as a char, or null.
    /// </summary>
    /// <returns>The cell's value as a char, or null.</returns>
    public abstract char? GetCharOrNull();

    /// <summary>
    /// Retrieves the cell's value as a boolean, or null.
    /// </summary>
    /// <returns>The cell's value as a boolean, or null.</returns>
    public abstract bool? GetBooleanOrNull();

    /// <summary>
    /// Retrieves the cell's value as a single-precision floating-point number, or null.
    /// </summary>
    /// <returns>The cell's value as a single-precision floating-point number, or null.</returns>
    public abstract float? GetSingleOrNull();

    /// <summary>
    /// Retrieves the cell's value as a double-precision floating-point number, or null.
    /// </summary>
    /// <returns>The cell's value as a double-precision floating-point number, or null.</returns>
    public abstract double? GetDoubleOrNull();

    /// <summary>
    /// Retrieves the cell's value as a decimal, or null.
    /// </summary>
    /// <returns>The cell's value as a decimal, or null.</returns>
    public abstract decimal? GetDecimalOrNull();

    /// <summary>
    /// Retrieves the cell's value as a 8-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 8-bit signed integer, or null.</returns>
    public abstract sbyte? GetSByteOrNull();

    /// <summary>
    /// Retrieves the cell's value as a 8-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 8-bit unsigned integer, or null.</returns>
    public abstract byte? GetByteOrNull();

    /// <summary>
    /// Retrieves the cell's value as a 16-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 16-bit signed integer, or null.</returns>
    public abstract short? GetInt16OrNull();

    /// <summary>
    /// Retrieves the cell's value as a 16-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 16-bit unsigned integer, or null.</returns>
    public abstract ushort? GetUInt16OrNull();

    /// <summary>
    /// Retrieves the cell's value as a 32-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 32-bit signed integer, or null.</returns>
    public abstract int? GetInt32OrNull();

    /// <summary>
    /// Retrieves the cell's value as a 32-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 32-bit unsigned integer, or null.</returns>
    public abstract uint? GetUInt32OrNull();

    /// <summary>
    /// Retrieves the cell's value as a 64-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 64-bit signed integer, or null.</returns>
    public abstract long? GetInt64OrNull();

    /// <summary>
    /// Retrieves the cell's value as a 64-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 64-bit unsigned integer, or null.</returns>
    public abstract ulong? GetUInt64OrNull();

    /// <summary>
    /// Retrieves the cell's value as a DateTime object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateTime object, or null.</returns>
    public abstract DateTime? GetDateTimeOrNull();

    /// <summary>
    /// Retrieves the cell's value as a DateOnly object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateOnly object, or null.</returns>
    public abstract DateOnly? GetDateOnlyOrNull();

    /// <summary>
    /// Retrieves the cell's value as a TimeOnly object, or null.
    /// </summary>
    /// <returns>The cell's value as a TimeOnly object, or null.</returns>
    public abstract TimeOnly? GetTimeOnlyOrNull();

    /// <summary>
    /// Retrieves the cell's value as a DateTimeOffset object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateTimeOffset object, or null.</returns>
    public abstract DateTimeOffset? GetDateTimeOffsetOrNull();
}

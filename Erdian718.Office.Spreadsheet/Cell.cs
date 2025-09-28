namespace Erdian718.Office.Spreadsheet;

/// <summary>
/// Represents a cell in a row.
/// </summary>
public class Cell
{
    /// <summary>
    /// Represents a empty cell.
    /// </summary>
    internal protected static Cell Empty { get; } = new();

    /// <summary>
    /// Gets a value indicating whether the cell is blank.
    /// </summary>
    public bool IsBlank
    {
        get
        {
            var value = Value;
            return value is string s ? string.IsNullOrEmpty(s) : value is null;
        }
    }

    /// <summary>
    /// Gets the underlying raw value stored in the cell.
    /// </summary>
    public virtual object? Value => null;

    /// <summary>
    /// Retrieves the cell's value as a string.
    /// </summary>
    /// <returns>The cell's value as a string.</returns>
    public virtual string GetString() => Value?.ToString() ?? string.Empty;

    /// <summary>
    /// Retrieves the cell's value as a char.
    /// </summary>
    /// <returns>The cell's value as a char.</returns>
    public virtual char GetChar() => Convert.ToChar(Value);

    /// <summary>
    /// Retrieves the cell's value as a boolean.
    /// </summary>
    /// <returns>The cell's value as a boolean.</returns>
    public virtual bool GetBoolean() => Convert.ToBoolean(Value);

    /// <summary>
    /// Retrieves the cell's value as a single-precision floating-point number.
    /// </summary>
    /// <returns>The cell's value as a single-precision floating-point number.</returns>
    public virtual float GetSingle() => Convert.ToSingle(Value);

    /// <summary>
    /// Retrieves the cell's value as a double-precision floating-point number.
    /// </summary>
    /// <returns>The cell's value as a double-precision floating-point number.</returns>
    public virtual double GetDouble() => Convert.ToDouble(Value);

    /// <summary>
    /// Retrieves the cell's value as a decimal.
    /// </summary>
    /// <returns>The cell's value as a decimal.</returns>
    public virtual decimal GetDecimal() => Convert.ToDecimal(Value);

    /// <summary>
    /// Retrieves the cell's value as a 8-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 8-bit signed integer.</returns>
    public virtual sbyte GetSByte() => Convert.ToSByte(Value);

    /// <summary>
    /// Retrieves the cell's value as a 8-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 8-bit unsigned integer.</returns>
    public virtual byte GetByte() => Convert.ToByte(Value);

    /// <summary>
    /// Retrieves the cell's value as a 16-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 16-bit signed integer.</returns>
    public virtual short GetInt16() => Convert.ToInt16(Value);

    /// <summary>
    /// Retrieves the cell's value as a 16-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 16-bit unsigned integer.</returns>
    public virtual ushort GetUInt16() => Convert.ToUInt16(Value);

    /// <summary>
    /// Retrieves the cell's value as a 32-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 32-bit signed integer.</returns>
    public virtual int GetInt32() => Convert.ToInt32(Value);

    /// <summary>
    /// Retrieves the cell's value as a 32-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 32-bit unsigned integer.</returns>
    public virtual uint GetUInt32() => Convert.ToUInt32(Value);

    /// <summary>
    /// Retrieves the cell's value as a 64-bit signed integer.
    /// </summary>
    /// <returns>The cell's value as a 64-bit signed integer.</returns>
    public virtual long GetInt64() => Convert.ToInt64(Value);

    /// <summary>
    /// Retrieves the cell's value as a 64-bit unsigned integer.
    /// </summary>
    /// <returns>The cell's value as a 64-bit unsigned integer.</returns>
    public virtual ulong GetUInt64() => Convert.ToUInt64(Value);

    /// <summary>
    /// Retrieves the cell's value as a DateTime object.
    /// </summary>
    /// <returns>The cell's value as a DateTime object.</returns>
    public virtual DateTime GetDateTime() => Convert.ToDateTime(Value);

    /// <summary>
    /// Retrieves the cell's value as a DateOnly object.
    /// </summary>
    /// <returns>The cell's value as a DateOnly object.</returns>
    public virtual DateOnly GetDateOnly() =>
        Value is DateOnly v ? v : DateOnly.FromDateTime(GetDateTime());

    /// <summary>
    /// Retrieves the cell's value as a TimeOnly object.
    /// </summary>
    /// <returns>The cell's value as a TimeOnly object.</returns>
    public virtual TimeOnly GetTimeOnly() =>
        Value is TimeOnly v ? v : TimeOnly.FromDateTime(GetDateTime());

    /// <summary>
    /// Retrieves the cell's value as a DateTimeOffset object.
    /// </summary>
    /// <returns>The cell's value as a DateTimeOffset object.</returns>
    public virtual DateTimeOffset GetDateTimeOffset() =>
        Value is DateTimeOffset v ? v : new(GetDateTime());

    /// <summary>
    /// Retrieves the cell's value as a string, or null.
    /// </summary>
    /// <returns>The cell's value as a string, or null.</returns>
    public string? GetStringOrNull()
    {
        try { return GetString(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a char, or null.
    /// </summary>
    /// <returns>The cell's value as a char, or null.</returns>
    public char? GetCharOrNull()
    {
        try { return GetChar(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a boolean, or null.
    /// </summary>
    /// <returns>The cell's value as a boolean, or null.</returns>
    public bool? GetBooleanOrNull()
    {
        try { return GetBoolean(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a single-precision floating-point number, or null.
    /// </summary>
    /// <returns>The cell's value as a single-precision floating-point number, or null.</returns>
    public float? GetSingleOrNull()
    {
        try { return GetSingle(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a double-precision floating-point number, or null.
    /// </summary>
    /// <returns>The cell's value as a double-precision floating-point number, or null.</returns>
    public double? GetDoubleOrNull()
    {
        try { return GetDouble(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a decimal, or null.
    /// </summary>
    /// <returns>The cell's value as a decimal, or null.</returns>
    public decimal? GetDecimalOrNull()
    {
        try { return GetDecimal(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 8-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 8-bit signed integer, or null.</returns>
    public sbyte? GetSByteOrNull()
    {
        try { return GetSByte(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 8-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 8-bit unsigned integer, or null.</returns>
    public byte? GetByteOrNull()
    {
        try { return GetByte(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 16-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 16-bit signed integer, or null.</returns>
    public short? GetInt16OrNull()
    {
        try { return GetInt16(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 16-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 16-bit unsigned integer, or null.</returns>
    public ushort? GetUInt16OrNull()
    {
        try { return GetUInt16(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 32-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 32-bit signed integer, or null.</returns>
    public int? GetInt32OrNull()
    {
        try { return GetInt32(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 32-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 32-bit unsigned integer, or null.</returns>
    public uint? GetUInt32OrNull()
    {
        try { return GetUInt32(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 64-bit signed integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 64-bit signed integer, or null.</returns>
    public long? GetInt64OrNull()
    {
        try { return GetInt64(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a 64-bit unsigned integer, or null.
    /// </summary>
    /// <returns>The cell's value as a 64-bit unsigned integer, or null.</returns>
    public ulong? GetUInt64OrNull()
    {
        try { return GetUInt64(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a DateTime object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateTime object, or null.</returns>
    public DateTime? GetDateTimeOrNull()
    {
        try { return GetDateTime(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a DateOnly object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateOnly object, or null.</returns>
    public DateOnly? GetDateOnlyOrNull()
    {
        try { return GetDateOnly(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a TimeOnly object, or null.
    /// </summary>
    /// <returns>The cell's value as a TimeOnly object, or null.</returns>
    public TimeOnly? GetTimeOnlyOrNull()
    {
        try { return GetTimeOnly(); } catch { return null; }
    }

    /// <summary>
    /// Retrieves the cell's value as a DateTimeOffset object, or null.
    /// </summary>
    /// <returns>The cell's value as a DateTimeOffset object, or null.</returns>
    public DateTimeOffset? GetDateTimeOffsetOrNull()
    {
        try { return GetDateTimeOffset(); } catch { return null; }
    }
}

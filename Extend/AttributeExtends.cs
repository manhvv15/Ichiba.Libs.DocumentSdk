using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using Ichiba.Libs.DocumentSdk.Constants;

namespace Ichiba.Libs.DocumentSdk.Extend;

/// <summary>
/// Represents a custom attribute that indicates a work sheet class.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public class WorkSheetAttribute : Attribute
{
    private readonly string _sheetName;
    private readonly int _headerRow;
    private readonly int _startCol;
    private readonly int _startRow;
    private readonly int _endCol;
    private readonly int _endRow;
    private readonly bool _isValidate;

    /// <summary>
    /// Initializes a new instance of the <see cref="WorkSheetAttribute"/> class.
    /// </summary>
    /// <param name="sheetName">The name of sheet to work.</param>
    /// <param name="headerRow">The header row to work.</param>
    /// <param name="startCol">The start column to work.</param>
    /// <param name="startRow">The start row to work.</param>
    /// <param name="endCol">The end column to work.</param>
    /// <param name="endRow">The end row to work.</param>
    /// <param name="isValidate">The flag to enable validate field.</param>
    public WorkSheetAttribute(string? sheetName = null, int headerRow = 0, int startCol = 0, int startRow = 0, int endCol = -1, int endRow = -1, bool isValidate = true)
        : base()
    {
        _sheetName = sheetName ?? CommonConstants.DefaultNameSheetGetValue;
        _headerRow = headerRow;
        _startCol = startCol;
        _startRow = startRow;
        _endCol = endCol;
        _endRow = endRow;
        _isValidate = isValidate;
    }

    public string GetSheetName()
    {
        return _sheetName;
    }

    public int GetHeaderRow()
    {
        return _headerRow >= 0 ? _headerRow : 0;
    }

    public int GetStartCol()
    {
        return _startCol >= 0 ? _startCol : 0;
    }

    public int GetStartRow()
    {
        return _startRow >= 0 ? _startRow : 0;
    }

    public int GetEndCol()
    {
        return _endCol;
    }

    public int GetEndRow()
    {
        return _endRow;
    }

    public bool IsValidate()
    {
        return _isValidate;
    }
}

[AttributeUsage(AttributeTargets.Property)]
public class DynamicAttribute : Attribute
{
}

[AttributeUsage(AttributeTargets.Property)]
public class IgnorePropertyAttribute : Attribute
{
}

/// <summary>
/// Represents a custom attribute that map value property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class AliasAttribute : Attribute
{
    public string _alias { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="AliasAttribute"/> class.
    /// </summary>
    /// <param name="alias">The name to map value.</param>
    public AliasAttribute(string alias)
    {
        _alias = alias;
    }

    public string GetAlias()
    {
        return _alias;
    }
}

/// <summary>
/// Represents a custom attribute that indicates a required property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class RequiredAttribute : ValidationAttribute
{
    /// <summary>
    /// Initializes a new instance of the <see cref="RequiredAttribute"/> class.
    /// </summary>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public RequiredAttribute(string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.Required) { }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value == null || string.IsNullOrEmpty(value.ToString()))
        {
            string memberName = validationContext.MemberName;
            //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
            //if (property != null)
            //{
            //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
            //    if (aliasAttribute != null)
            //    {
            //        memberName = aliasAttribute.GetAlias();
            //    }
            //}
            return new ValidationResult(ErrorMessage, new[] { memberName });
        }
        return ValidationResult.Success;
    }
}

/// <summary>
/// Represents a custom attribute that validates a date property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class DateAttribute : ValidationAttribute
{
    private readonly string _dateFormat;

    /// <summary>
    /// Initializes a new instance of the <see cref="DateAttribute"/> class.
    /// </summary>
    /// <param name="dateFormat">The format of the date.</param>
    /// <param name="errorMessage">The error message to display if the date is invalid.</param>
    public DateAttribute(string? dateFormat = null, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidDateFormat)
    {
        _dateFormat = dateFormat ?? CommonConstants.DefaultDateFormat;
    }

    /// <summary>
    /// Gets the format of the date.
    /// </summary>
    /// <returns>The format of the date.</returns>
    public string GetDateFormat()
    {
        return _dateFormat;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        return ValidationResult.Success;
    }
}

/// <summary>
/// Represents a custom attribute that validates a datetime property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class DateTimeAttribute : ValidationAttribute
{
    private readonly string _dateTimeFormat;

    /// <summary>
    /// Initializes a new instance of the <see cref="DateTimeAttribute"/> class.
    /// </summary>
    /// <param name="dateTimeFormat">The format of the datetime.</param>
    /// <param name="errorMessage">The error message to display if the date is invalid.</param>
    public DateTimeAttribute(string? dateTimeFormat = null, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidDateTimeFormat)
    {
        _dateTimeFormat = dateTimeFormat ?? CommonConstants.DefaultDateTimeFormat;
    }

    public string GetDateTimeFormat()
    {
        return _dateTimeFormat;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        return ValidationResult.Success;
        //if (value != null)
        //{
        //    var dateString = value.ToString();
        //    if (DateTime.TryParseExact(dateString, _dateTimeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
        //    {
        //        return ValidationResult.Success;
        //    }
        //}

        //string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        //return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates a float property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class FloatAttribute : ValidationAttribute
{
    /// <summary>
    /// Initializes a new instance of the <see cref="FloatAttribute"/> class.
    /// </summary>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public FloatAttribute(string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidFloatPrecision) { }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && float.TryParse(value.ToString(), out float _))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates a float precision property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class FloatPrecisionAttribute : ValidationAttribute
{
    private readonly int _precision;

    /// <summary>
    /// Initializes a new instance of the <see cref="FloatPrecisionAttribute"/> class.
    /// </summary>
    /// <param name="precision">The precision of float number</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public FloatPrecisionAttribute(int precision, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidFloatPrecision)
    {
        _precision = precision;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && float.TryParse(value.ToString(), out float floatValue))
        {
            var decimalPlaces = BitConverter.GetBytes(decimal.GetBits((decimal)floatValue)[3])[2];
            if (decimalPlaces <= _precision)
            {
                return ValidationResult.Success;
            }
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates match regex property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class MatchAttribute : ValidationAttribute
{
    private readonly string _pattern;

    /// <summary>
    /// Initializes a new instance of the <see cref="MatchAttribute"/> class.
    /// </summary>
    /// <param name="pattern">The pattern to match</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public MatchAttribute(string pattern, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidFormat)
    {
        _pattern = pattern;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && Regex.IsMatch(value.ToString(), _pattern))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates email property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class EmailAttribute : ValidationAttribute
{
    private readonly string _pattern;

    /// <summary>
    /// Initializes a new instance of the <see cref="EmailAttribute"/> class.
    /// </summary>
    /// <param name="pattern">The pattern to match</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public EmailAttribute(string? pattern = null, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidEmailFormat)
    {
        _pattern = pattern ?? CommonConstants.DefaultEmailFormat;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && Regex.IsMatch(value.ToString(), _pattern))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates phonenumber property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class PhoneNumberAttribute : ValidationAttribute
{
    private readonly string _pattern;

    /// <summary>
    /// Initializes a new instance of the <see cref="PhoneNumberAttribute"/> class.
    /// </summary>
    /// <param name="pattern">The pattern to match</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public PhoneNumberAttribute(string? pattern = null, string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidPhoneNumberFormat)
    {
        _pattern = pattern ?? CommonConstants.DefaultPhoneFormat;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && Regex.IsMatch(value.ToString(), _pattern))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates long property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class LongAttribute : ValidationAttribute
{
    /// <summary>
    /// Initializes a new instance of the <see cref="LongAttribute"/> class.
    /// </summary>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public LongAttribute(string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidNumericValue) { }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && long.TryParse(value.ToString(), out _))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates long property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class IntAttribute : ValidationAttribute
{
    /// <summary>
    /// Initializes a new instance of the <see cref="IntAttribute"/> class.
    /// </summary>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public IntAttribute(string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.InvalidNumericValue) { }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && int.TryParse(value.ToString(), out _))
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates unique property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class UniqueValuesAttribute : ValidationAttribute
{
    /// <summary>
    /// Initializes a new instance of the <see cref="UniqueValuesAttribute"/> class.
    /// </summary>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public UniqueValuesAttribute(string? errorMessage = null) : base(errorMessage ?? ErrorMessageConstants.Unique) { }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value is IEnumerable<object> enumerable)
        {
            string memberName = validationContext.MemberName;
            //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
            //if (property != null)
            //{
            //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
            //    if (aliasAttribute != null)
            //    {
            //        memberName = aliasAttribute.GetAlias();
            //    }
            //}

            var set = new HashSet<object>();
            foreach (var item in enumerable)
            {
                if (!set.Add(item))
                {
                    return new ValidationResult(ErrorMessage, new[] { memberName });
                }
            }
        }
        return ValidationResult.Success;
        //return new ValidationResult("Invalid data type", new[] { validationContext.MemberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates min length property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class MinLengthAttribute : ValidationAttribute
{
    private int _minLength;

    /// <summary>
    /// Initializes a new instance of the <see cref="MinLengthAttribute"/> class.
    /// </summary>
    /// <param name="minLength">The number to check min length.</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public MinLengthAttribute(int minLength, string? errorMessage = null) : base(errorMessage ?? string.Format(ErrorMessageConstants.MinLength, minLength))
    {
        _minLength = minLength;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && value.ToString().Length >= _minLength)
        {
            return ValidationResult.Success;
        }

        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

/// <summary>
/// Represents a custom attribute that validates max length property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class MaxLengthAttribute : ValidationAttribute
{
    private int _maxLength;

    /// <summary>
    /// Initializes a new instance of the <see cref="MaxLengthAttribute"/> class.
    /// </summary>
    /// <param name="maxLength">The number to check max length.</param>
    /// <param name="errorMessage">The error message to display if the property is not provided.</param>
    public MaxLengthAttribute(int maxLength, string? errorMessage = null) : base(errorMessage ?? string.Format(ErrorMessageConstants.MaxLength, maxLength))
    {
        _maxLength = maxLength;
    }

    protected override ValidationResult IsValid(object value, ValidationContext validationContext)
    {
        if (value != null && value.ToString().Length <= _maxLength)
        {
            return ValidationResult.Success;
        }
        string memberName = validationContext.MemberName;
        //var property = validationContext.ObjectType.GetProperty(validationContext.MemberName);
        //if (property != null)
        //{
        //    var aliasAttribute = property.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute;
        //    if (aliasAttribute != null)
        //    {
        //        memberName = aliasAttribute.GetAlias();
        //    }
        //}
        return new ValidationResult(ErrorMessage, new[] { memberName });
    }
}

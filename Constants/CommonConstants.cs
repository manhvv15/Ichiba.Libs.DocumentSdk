namespace Ichiba.Libs.DocumentSdk.Constants;

public static class CommonConstants
{
    public const string DefaultEmailFormat = "^(?:[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+(?:\\.[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:[\\x01-\\x08\\x0b\\x0c\\x0e-\\x1f\\x21\\x23-\\x5b\\x5d-\\x7f]|\\\\[\\x01-\\x09\\x0b\\x0c\\x0e-\\x7f])*\")@(?:(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\\.)+[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?|\\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-zA-Z0-9-]*[a-zA-Z0-9]:(?:[\\x01-\\x08\\x0b\\x0c\\x0e-\\x1f\\x21-\\x5a\\x53-\\x7f]|\\\\[\\x01-\\x09\\x0b\\x0c\\x0e-\\x7f])+)\\])\\S+";
    public const string DefaultPhoneFormat = @"^\+?[0-9][0-9]{3,25}$";
    public const string DefaultDateTimeFormat = "dd/MM/yyyy HH:mm:ss";
    public const string DefaultDateFormat = "dd/MM/yyyy";
    public const string DefaultNameSheetGetValue = "Sheet1";
    public const string KeyProtected = "724626c32d37f9b78c22793bc1fb802748f80cb5ce4b58cedefbee3914271d66";
    public const string NameSheetKey = "Key";
    public const string PasswordSheetKey = "f2f9068b283f0c9a22d0326c27769bc79cff9dfd15ed147d47bc4ad91d270fd1";
    public const string CellContainKey = "A2";
}

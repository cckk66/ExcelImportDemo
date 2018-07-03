using System;

namespace ExcelImport.Atrribute.ValueAttribute
{
    public class EnumConvertAttribute : ConvertValueAttribute
    {
        public Type enumType { get; set; }
        public EnumConvertAttribute(Type _enumType)
        {
            enumType = _enumType;
        }
        public override object GetConvertValueValue(object value)
        {
            try
            {
                if (value != DBNull.Value && value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    return System.Enum.Format(enumType, System.Enum.Parse(enumType, value.ToString()), "D");
                }
                return value;
            }
            catch (Exception)
            {

                return value;
            }
            
        }
    }
}

using System;

namespace ExcelImport.Atrribute.ValueAttribute
{
    public class EnumConvertAttribute : ConvertValueAttribute
    {
        public Type enumType { get; set; }
        public EnumConvertAttribute(Type _enumType) {
            enumType = _enumType;
        }
        public override object GetConvertValueValue(object value)
        {
            return System.Enum.Format(enumType, System.Enum.Parse(enumType, "女"), "D");
        }
    }
}

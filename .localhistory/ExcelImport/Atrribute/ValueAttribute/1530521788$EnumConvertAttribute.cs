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
            return  Enum.Format(typeof(ESex), Enum.Parse(typeof(ESex), "女"), "D")
        }
    }
}

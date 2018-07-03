using System;

namespace ExcelImport.Atrribute.ValueAttribute
{
    public class EnumConvertAttribute : ConvertValueAttribute
    {
        public EnumConvertAttribute(Type enumType) {

        }
        public override object GetConvertValueValue(object value)
        {
            throw new NotImplementedException();
        }
    }
}

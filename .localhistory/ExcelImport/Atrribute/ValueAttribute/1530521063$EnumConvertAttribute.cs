using System;

namespace ExcelImport.Atrribute.ValueAttribute
{
    public class EnumConvertAttribute : ConvertValueAttribute
    {
        public override object GetConvertValueValue(object value)
        {
            if (value != DBNull.Value && value != null && !string.IsNullOrEmpty(value.ToString()))
            {
                ESex esex = (ESex)Enum.Parse(typeof(ESex), value.ToString());
                return Enum.ToObject(typeof(ESex), esex);

            }
            return value;
        }
    }
}

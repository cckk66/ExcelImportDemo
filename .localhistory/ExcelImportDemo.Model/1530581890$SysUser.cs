using ExcelImport.Atrribute;
using ExcelImport.Atrribute.ValueAttribute;
using ExcelImport.Interface;
using System;

namespace ExcelImportDemo.Model
{

    public class SysUser : IExeclImport
    {
        [ExeclMapping("姓名")]
        public string Name { get; set; }
        [ExeclMapping("年龄")]
        public int Age { get; set; }
        [ExeclMapping("地址")]
        public string Address { get; set; }
        [ExeclMapping("性别")]
        [EnumConvert(typeof(ESex))]
        public int Sex { get; set; }
        public string ExeclRowErrorMsg { get; set; }
    }

    public enum ESex
    {
        男 = 1,
        女 = 2
    }
    public class ESexAppointValueAttribute : ConvertValueAttribute
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

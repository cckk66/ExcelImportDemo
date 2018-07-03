using ExcelImport.Atrribute;
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
        [ESexAppointValue]
        public int Sex { get; set; }
    }

    public enum ESex
    {
        男 = 1,
        女 = 2
    }
    public class ESexAppointValueAttribute : AppointValueAttribute
    {
        public override object GetAppointValue(object value)
        {

            if (value != DBNull.Value && value != null)
            {
                if (value.ToString() == "男")
                {

                }
            }
            return value;
        }
    }
}

using System;

namespace ExcelImport.Atrribute
{
    /// <summary>
    /// 转换值
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class ConvertValueAttribute : Attribute
    {
        public abstract object GetAppointValue(object value);
    }
}

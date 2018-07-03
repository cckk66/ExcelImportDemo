using System;

namespace ExcelImport.Atrribute
{
    /// <summary>
    /// 转换值基类 可继承该类指定值
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class ConvertValueAttribute : Attribute
    {
        public abstract object GetAppointValue(object value);
    }
}

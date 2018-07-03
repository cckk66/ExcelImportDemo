using System;

namespace ExcelImport.Atrribute
{
    /// <summary>
    /// 转换值基类 可继承该类指定值
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class ConvertValueAttribute<T> : Attribute
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value">excel值</param>
        /// <returns>转换值</returns>
        public abstract object GetConvertValueValue(object value);
    }
}


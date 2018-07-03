using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    /// <summary>
    /// 验证父类，自定义继承扩展
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class ExcelBaseValidationAttribute : Attribute
    {
        public ExcelBaseValidationAttribute(string _ErrorMsg)
        {
            ErrorMsg = _ErrorMsg;
        }
        public string ErrorMsg { get; set; }
        public abstract bool IsValid(object value);
    }
}

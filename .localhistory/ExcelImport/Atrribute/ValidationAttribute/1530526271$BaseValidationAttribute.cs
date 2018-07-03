using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class BaseValidationAttribute : Attribute
    {
        public BaseValidationAttribute(string _ErrorMsg) {
            ErrorMsg = _ErrorMsg;
        }
        public  virtual string ErrorMsg { get; set; }
        public abstract bool IsValid(object value);
    }
}

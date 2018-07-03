using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class BaseValidationAttribute : Attribute
    {
        public abstract string ErrorMsg { get; set; }
        public abstract bool IsValid(object value);
    }
}

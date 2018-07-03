using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class BaseValidationAttribute : Attribute
    {
        public abstract bool IsValid(object value)
    }
}

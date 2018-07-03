using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class CustomValidationAttribute : Attribute
    {
        public virtual string ErrorMsg { get; set; }

        public string ValidateRule { get; set; }
        public CustomValidationAttribute(string validateRule, string errorMsg)
        {
            ErrorMsg = errorMsg;
            ValidateRule = validateRule;
        }
        /// <summary>
        /// 自定义验证规则
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public virtual bool IsValid(object value)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(value.ToString(), ValidateRule);
        }
    }
}

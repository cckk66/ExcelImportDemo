﻿namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class MyCustomValidationAttribute : Attribute
    {
        public override string EasyUIVtName { get; set; } = "Custom";
        public override string ErrorMsg { get; set; }

        public string ValidateRule { get; set; }
        public MyCustomValidationAttribute(string validateRule, string errorMsg)
        {
            ErrorMsg = errorMsg;
            ValidateRule = validateRule;
            EasyUIVtName = $"Custom[\"{validateRule}\",\"{errorMsg}\"]";
        }
        /// <summary>
        /// 自定义验证规则
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override bool IsValid(object value)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(value.ToString(), ValidateRule);
        }
    }
}

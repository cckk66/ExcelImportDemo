namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class CustomValidationAttribute : BaseValidationAttribute
    {
      

        public string ValidateRule { get; set; }
        public override string ErrorMsg { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

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
        public override bool IsValid(object value)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(value.ToString(), ValidateRule);
        }
    }
}

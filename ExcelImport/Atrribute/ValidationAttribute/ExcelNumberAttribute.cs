using System.Text.RegularExpressions;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class ExcelNumberAttribute : ExcelBaseValidationAttribute
    {
        public ExcelNumberAttribute(string _ErrorMsg) : base(_ErrorMsg)
        {
            ErrorMsg = _ErrorMsg;
        }
    
        public override bool IsValid(object value)
        {
            return Regex.IsMatch(value.ToString(), @"^-?\d+\.?\d*$");
        }
    }
}

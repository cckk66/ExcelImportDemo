namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class RequiredValidationAttribute : BaseValidationAttribute
    {
        public override string ErrorMsg { get; set; }
        public override bool IsValid(object value)
        {
            throw new System.NotImplementedException();
        }
    }
}

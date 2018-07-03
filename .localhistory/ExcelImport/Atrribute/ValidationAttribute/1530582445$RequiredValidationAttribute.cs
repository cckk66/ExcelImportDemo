namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class RequiredValidationAttribute : BaseValidationAttribute
    {
        public RequiredValidationAttribute(string _ErrorMsg) : base(_ErrorMsg)
        {
        }


        public override bool IsValid(object value)
        {
            throw new System.NotImplementedException();
        }
    }
}

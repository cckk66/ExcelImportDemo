namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class MyLengthAttribute : CustomValidationAttribute
    {
        public override string ErrorMsg { get; set; }
        private int max;
        private int min = -1;

        public MyLengthAttribute(int max)
        {
            ErrorMsg = $"长度最大为{max}";
            this.max = max;
        }
        public MyLengthAttribute(int max, int min)
        {
            ErrorMsg = $"长度应该应该为{min}和{max}之间！";
            this.max = max;
            this.min = min;
        }
        public override bool IsValid(object value)
        {
            if (value.ToString().Length < min || value.ToString().Length > max)
            { return false; }
            else
            { return true; }
        }
    }
}

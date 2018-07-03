namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class MyLengthAttribute : MyValidationAttribute
    {
        public override string EasyUIVtName { get; set; }
        public override string ErrorMsg { get; set; }
        private int max;
        private int min = -1;

        public MyLengthAttribute(int max)
        {
            EasyUIVtName = $"MyLength[0,{max}]";
            ErrorMsg = $"长度最大为{max}";
            this.max = max;
        }
        public MyLengthAttribute(int max, int min)
        {
            EasyUIVtName = $"MyLength[{min},{max}]";
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

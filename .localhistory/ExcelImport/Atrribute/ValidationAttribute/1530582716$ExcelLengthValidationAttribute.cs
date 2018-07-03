namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class MyLengthAttribute : ExcelBaseValidationAttribute
    {
        private int max;
        private int min = -1;
        public MyLengthAttribute(int max, string _ErrorMsg) : base(_ErrorMsg)
        {
            if (string.IsNullOrEmpty(_ErrorMsg))
                ErrorMsg = $"长度最大为{max}";
            else
                ErrorMsg = _ErrorMsg;
          this.max = max;
        }


        public MyLengthAttribute(int max, int min, string _ErrorMsg) : base(_ErrorMsg)
        {
            if (string.IsNullOrEmpty(_ErrorMsg))
                ErrorMsg = $"长度应该应该为{min}和{max}之间！";
            else
                ErrorMsg = _ErrorMsg;
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

namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class ExcelNumberAttribute : ExcelBaseValidationAttribute
    {
        public ExcelNumberAttribute(string _ErrorMsg) : base(_ErrorMsg)
        {
        }

        //    public override string EasyUIVtName { get; set; } = "Number";
        //    public override string ErrorMsg { get; set; } = "只能输入数字";

        //    public MyNumberAttribute()
        //    {
        //    }
        //    public MyNumberAttribute(string easyUIVtName)
        //    {
        //        EasyUIVtName = easyUIVtName;
        //    }
        //    /// <summary>
        //    /// 后台验证规则
        //    /// </summary>
        //    /// <param name="value"></param>
        //    /// <returns></returns>
        //    public override bool IsValid(object value)
        //    {
        //        return Regex.IsMatch(value.ToString(), @"^-?\d+\.?\d*$");
        //    }
        //{
        public override bool IsValid(object value)
        {
            throw new System.NotImplementedException();
        }
    }
}

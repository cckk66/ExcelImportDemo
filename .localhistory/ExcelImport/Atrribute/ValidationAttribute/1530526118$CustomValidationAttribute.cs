﻿using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class CustomValidationAttribute : BaseValidationAttribute
    {


        public override string ErrorMsg { get; set; }
        public CustomValidationAttribute(string _ErrorMsg)
        {
            ErrorMsg = _ErrorMsg
        }

        /// <summary>
        /// 自定义验证规则
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override bool IsValid(object value)
        {
            return DBNull.Value == value && value != null;
        }
    }
}

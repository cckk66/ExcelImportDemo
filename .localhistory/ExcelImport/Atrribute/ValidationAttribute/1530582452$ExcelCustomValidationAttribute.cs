﻿using System;

namespace ExcelImport.Atrribute.ValidationAttribute
{
    public class CustomValidationAttribute : ExcelBaseValidationAttribute
    {

        public CustomValidationAttribute(string _ErrorMsg) : base(_ErrorMsg)
        {
            ErrorMsg = _ErrorMsg;
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

﻿using System;

namespace ExcelImport.Atrribute.ValueAttribute
{
    /// <summary>
    /// 制定枚举转换
    /// </summary>
    public enum EnumConvertValue
    {
        Name,
        Value
    }
    public class EnumConvertAttribute : ConvertValueAttribute
    {
        public Type enumType { get; set; }
        public EnumConvertValue enumConvertValue { get; set; } = EnumConvertValue.Value;
        public EnumConvertAttribute(Type _enumType)
        {
            enumType = _enumType;
        }
        public EnumConvertAttribute(Type _enumType)
        {
            enumType = _enumType;
        }
        public override object GetConvertValueValue(object value)
        {
            try
            {
                if (value != DBNull.Value && value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    return System.Enum.Format(enumType, System.Enum.Parse(enumType, value.ToString()), "D");
                }
                return value;
            }
            catch (Exception ex)
            {

                return value;
            }

        }
    }
}

using System;

namespace ExcelImport.Common
{
    public static class ObjectExtend
    {
        public static string ToDefaultString(this object obj, string defaultValue)
        {
            return (obj == DBNull.Value) ? defaultValue : obj.ToString();
        }
    }
}

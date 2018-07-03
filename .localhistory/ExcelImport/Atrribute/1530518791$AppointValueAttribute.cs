using System;

namespace ExcelImport.Atrribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class AppointValueAttribute : Attribute
    {
        public abstract object GetAppointValue(object value);
    }
}

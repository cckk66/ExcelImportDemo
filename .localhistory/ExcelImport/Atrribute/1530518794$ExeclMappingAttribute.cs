using System;

namespace ExcelImport.Atrribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExeclMapping : System.Attribute
    {
        public string ExeclHeadName { get; set; }
        public ExeclMapping(string _ExeclHeadName)
        {
            ExeclHeadName = _ExeclHeadName;
        }
    }
}

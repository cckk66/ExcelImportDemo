using System;

namespace ExcelImport.Atrribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExeclMappingAttribute : System.Attribute
    {
        public string ExeclHeadName { get; set; }
        public ExeclMappingAttribute(string _ExeclHeadName)
        {
            ExeclHeadName = _ExeclHeadName;
        }
    }
}

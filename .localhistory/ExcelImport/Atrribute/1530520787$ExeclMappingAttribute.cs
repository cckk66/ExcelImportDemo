using System;

namespace ExcelImport.Atrribute
{
    /// <summary>
    /// 映射
    /// </summary>
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

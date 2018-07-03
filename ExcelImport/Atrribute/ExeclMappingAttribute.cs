using System;

namespace ExcelImport.Atrribute
{
    /// <summary>
    /// 映射execl表头与对应类名称
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

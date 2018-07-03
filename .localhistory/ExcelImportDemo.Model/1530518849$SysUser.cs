using ExcelImport.Atrribute;
using ExcelImport.Interface;

namespace ExcelImportDemo.Model
{

    public class SysUser: IExeclImport
    {
        [ExeclMappingAttribute("姓名")]
        public string Name { get; set; }
        [ExeclMappingAttribute("年龄")]
        public int Age { get; set; }
        [ExeclMappingAttribute("地址")]
        public string Address { get; set; }
    }
}

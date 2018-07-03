using ExcelImport.Atrribute;
using ExcelImport.Interface;

namespace ExcelImportDemo.Model
{

    public class SysUser: IExeclImport
    {
        [ExeclMapping("姓名")]
        public string Name { get; set; }
        [ExeclMapping("年龄")]
        public int Age { get; set; }
        [ExeclMapping("地址")]
        public string Address { get; set; }
        [ExeclMapping("性别")]
        public int Sex { get; set; }
    }
}

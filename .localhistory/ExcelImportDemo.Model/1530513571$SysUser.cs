﻿using ExcelImport.Atrribute;

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
    }
}

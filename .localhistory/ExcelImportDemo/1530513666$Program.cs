using ExcelImport;
using ExcelImportDemo.Model;
using System.Collections.Generic;

namespace ExcelImportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelImportHelper excelImportHelper = new ExcelImportHelper();
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
        }
    }
}

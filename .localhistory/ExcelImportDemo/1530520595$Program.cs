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
            System.Console.WriteLine($"读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(account); Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

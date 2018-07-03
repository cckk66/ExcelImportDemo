using ExcelImport;
using ExcelImportDemo.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace ExcelImportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (object o in Enum.GetValues(typeof(ESex))) {
                Console.WriteLine("{0}:{1}", o, Enum.Format(typeof(ESex), o, "D"));
            }

            ExcelImportHelper excelImportHelper = new ExcelImportHelper();
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            System.Console.WriteLine($"读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

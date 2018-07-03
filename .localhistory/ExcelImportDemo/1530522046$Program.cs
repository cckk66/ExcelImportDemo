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
            Array array= Enum.GetValues(typeof(ESex));



            Console.WriteLine( Enum.Format(typeof(ESex), Enum.Parse(typeof(ESex), "女"), "D"));
            Console.WriteLine(Enum.Format(typeof(ESex), Enum.Parse(typeof(ESex), "女"), "D"));
            ExcelImportHelper excelImportHelper = new ExcelImportHelper();
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            System.Console.WriteLine($"读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

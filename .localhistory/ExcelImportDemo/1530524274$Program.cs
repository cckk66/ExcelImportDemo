using ExcelImport;
using ExcelImportDemo.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelImportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //Array array= Enum.GetValues(typeof(ESex));


            //Console.WriteLine(System.Enum.GetName(typeof(ESex), 1));
            //Console.WriteLine(System.Enum.GetName(typeof(ESex), 2));
            //Console.WriteLine( Enum.Format(typeof(ESex), Enum.Parse(typeof(ESex), "女"), "D"));
            //Console.WriteLine(Enum.Format(typeof(ESex), Enum.Parse(typeof(ESex), "女"), "D"));
            ExcelImportHelper excelImportHelper = new ExcelImportHelper();
            System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
            stopwatch.Stop(); //  停止监视
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            TimeSpan timespan = stopwatch.Elapsed; //  获取当前实例测量得出的总时间
            System.Console.WriteLine($"第一次转换 {sysUserList.Count}");
            System.Console.WriteLine($"读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

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
            stopwatch.Start(); //  开始监视代码运行时间
          
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            stopwatch.Stop(); //  停止监视
            TimeSpan timespan = stopwatch.Elapsed; //  获取当前实例测量得出的总时间
            System.Console.WriteLine($"第一次转换 {timespan.TotalSeconds}");
            System.Console.WriteLine($"第一次读取execl {sysUserList.Count}");
            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            stopwatch.Reset();

            stopwatch.Start(); //  开始监视代码运行时间

            sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            stopwatch.Stop(); //  停止监视
             timespan = stopwatch.Elapsed; //  获取当前实例测量得出的总时间
            System.Console.WriteLine($"第二次转换 {timespan.TotalSeconds}");
            System.Console.WriteLine($"第二次读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

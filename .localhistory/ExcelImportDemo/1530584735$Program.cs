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

            //1.获取模块的完整路径。  
            string path1 = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

            //2.获取和设置当前目录(该进程从中启动的目录)的完全限定目录  
            string path2 = System.Environment.CurrentDirectory;

            //3.获取应用程序的当前工作目录  
            string path3 = System.IO.Directory.GetCurrentDirectory();

            //4.获取程序的基目录  
            string path4 = System.AppDomain.CurrentDomain.BaseDirectory;

            //5.获取和设置包括该应用程序的目录的名称  
            string path5 = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            //6.获取启动了应用程序的可执行文件的路径  
            string path6 = System.Windows.Forms.Application.StartupPath;

            //7.获取启动了应用程序的可执行文件的路径及文件名  
            string path7 = System.Windows.Forms.Application.ExecutablePath;

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

            json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

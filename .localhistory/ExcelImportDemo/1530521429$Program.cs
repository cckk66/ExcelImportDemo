﻿using ExcelImport;
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

            Console.WriteLine(Enum.ToObject(typeof(ESex),(ESex) Enum.Parse(typeof(ESex), "男")));
            Console.WriteLine(Enum.ToObject(typeof(ESex), (ESex)Enum.Parse(typeof(ESex), "女")));

            ExcelImportHelper excelImportHelper = new ExcelImportHelper();
            List<SysUser> sysUserList = excelImportHelper.DataTableToList<SysUser>(@"C:\Users\chenke\Desktop\execl测试.xlsx");
            System.Console.WriteLine($"读取execl {sysUserList.Count}");

            string json = JsonConvert.SerializeObject(sysUserList);
            Console.WriteLine(json);
            System.Console.Read();
        }
    }
}

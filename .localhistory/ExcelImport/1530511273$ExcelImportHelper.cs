﻿using ExcelImportDemo;
using System.Data;

namespace ExcelImport
{
    public abstract class ExcelImportHelper
    {
        private static Office07Helper office07Helper = new Office07Helper();
        public int MaxLeng { get; set; } = 500;
        public string ErrorMsg { get; set; } = "";

        protected DataTable ExeclToDataTable(string filePath, string sheetName)
        {
            return office07Helper.ReaderExcel(filePath, sheetName);
        }
    }
}

using ExcelImport.Interface;
using ExcelImportDemo;
using System.Collections.Generic;
using System.Data;

namespace ExcelImport
{
    public abstract class ExcelImportHelper
    {
        private static Office07Helper office07Helper = new Office07Helper();
        public int MaxLeng { get; set; } = 500;
        public string ErrorMsg { get; set; } = "";

        protected DataTable ExeclToDataTable(string filePath) => this.ExeclToDataTable(filePath, "");

        protected DataTable ExeclToDataTable(string filePath, string sheetName)
        {
            return office07Helper.ReaderExcel(filePath, sheetName);
        }
        public List<T> DataTableToList<T>(string filePath, string sheetName) where T : IExeclImport
        {
            List<T> listEntity = new List<T>();

            return listEntity;
        }
    }
}

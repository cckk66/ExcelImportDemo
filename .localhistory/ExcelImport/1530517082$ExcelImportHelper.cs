using ExcelImport.Atrribute;
using ExcelImport.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace ExcelImport
{
    public class ExcelImportHelper
    {
        private static ExcelReader excelReader = new ExcelReader();
        
        public int MaxLeng { get; set; } = 500;
        public string ErrorMsg { get; set; } = "";

      


        public List<T> DataTableToList<T>(string filePath) where T : IExeclImport => this.DataTableToList<T>(filePath, "");
        public List<T> DataTableToList<T>(string filePath, string sheetName) where T : IExeclImport
        {
            List<T> listEntity = new List<T>();
            (string[] headArr, DataTable dt) ExcelInfo = excelReader.ReaderExcel(filePath, sheetName);
            if (ExcelInfo==null)
            {

            }
            DataTable dt = ExcelInfo.dt;
            if (dt.Rows.Count > this.MaxLeng)
            {
                this.ErrorMsg = $"超过最大条数{MaxLeng}";
                return null;
            }

            var PropertyInfos = typeof(T).GetProperties().Where(m => m.GetCustomAttributes(typeof(ExeclMapping), true) != null);
            if (PropertyInfos == null || PropertyInfos.Count() <= 0)
            {
                this.ErrorMsg = $"传入实体无Execl映射信息";
                return null;
            }

            foreach (DataRow row in dt.Rows)
            {
                T t = default(T);
                foreach (PropertyInfo Prty in PropertyInfos)
                {
                    if (dt.Columns.Contains(Prty.Name))
                    {
                        // 判断此属性是否有Setter
                        if (!Prty.CanWrite) continue;//该属性不可写，直接跳出
                                                     //取值  
                        object value = row[Prty.Name];
                        //如果非空，则赋给对象的属性  
                        if (value != DBNull.Value)
                            Prty.SetValue(t, value, null);

                    }
                }
                listEntity.Add(t);
            }
            return listEntity;
        }
    }
}

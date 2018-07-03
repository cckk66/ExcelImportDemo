using ExcelImport.Atrribute;
using ExcelImport.Common;
using ExcelImport.Interface;
using ExcelImportDemo;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelImport
{
    public class ExcelImportHelper
    {
        private static Office07Helper office07Helper = new Office07Helper();
        public int MaxLeng { get; set; } = 500;
        public string ErrorMsg { get; set; } = "";

        public List<T> DataTableToList<T>(string filePath) where T : IExeclImport => this.DataTableToList<T>(filePath, "");

        /// <summary>
        /// 读取标准Excel格式
        /// </summary>
        /// <param name="filepath">文件地址</param>
        /// <param name="sheetName">工作博名称</param>
        /// <param name="ReaderRow">从第几行开始</param>
        /// <param name="haveNumber">是否需要计数字段</param>
        /// <returns>DataTable数据集合</returns>
        private DataTable ReaderExcel(string filepath, string sheetName)
        {

            DataTable dt = new DataTable();
            using (FileStream fs = File.OpenRead(filepath))
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);

                //如果ReaderRow大于0,则不是从一行开始读取.创建默认表头
                ISheet sheet = string.IsNullOrWhiteSpace(sheetName) ? wk.GetSheetAt(0) : wk.GetSheet(sheetName);


                for (int i = 0; i < sheet.GetRow(0).LastCellNum; i++)
                {
                    dt.Columns.Add(new DataColumn(sheet.GetRow(0).GetCell(i).ToString()));
                }

                #region 读取入内容



                for (int i = ReaderRow; i <= sheet.LastRowNum; i++)
                {
                    DataRow row = dt.NewRow();
                    for (int j = 0; j < sheet.GetRow(0).LastCellNum; j++)
                    {
                        row[j] = sheet.GetRow(i).GetCell(j).ToDefaultString();
                    }
                    dt.Rows.Add(row);
                }

                #endregion
            }
            return dt;
        }
        public List<T> DataTableToList<T>(string filePath, string sheetName) where T : IExeclImport
        {
            List<T> listEntity = new List<T>();
            DataTable dt = office07Helper.ReaderExcel(filePath, sheetName);
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

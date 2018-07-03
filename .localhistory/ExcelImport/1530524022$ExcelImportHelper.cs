using ExcelImport.Atrribute;
using ExcelImport.Cache;
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




        public List<T> DataTableToList<T>(string filePath) where T : IExeclImport, new() => this.DataTableToList<T>(filePath, "");
        public virtual List<T> DataTableToList<T>(string filePath, string sheetName) where T : IExeclImport, new()
        {
            List<T> listEntity = new List<T>();
            DataTable dt = excelReader.ReaderExcel(filePath, sheetName);
            if (dt.Rows.Count > this.MaxLeng)
            {
                this.ErrorMsg = $"超过最大条数{MaxLeng}";
                return null;
            }
            string TName = GetTName<T>();

            #region 获取实体类映射集合  
            string ExeclMappingPropertyInfosKey = $"{TName}_ExeclMappingPropertys";
            IEnumerable<PropertyInfo> PropertyInfos = CacheHelper.GetCache<IEnumerable<PropertyInfo>>(ExeclMappingPropertyInfosKey);
            if (PropertyInfos == null)
            {
                PropertyInfos = typeof(T).GetProperties().Where(m => m.GetCustomAttributes(typeof(ExeclMappingAttribute), true) != null);
                CacheHelper.SetCache(ExeclMappingPropertyInfosKey, PropertyInfos);
            }
            #endregion

            if (PropertyInfos == null || PropertyInfos.Count() <= 0)
            {
                this.ErrorMsg = $"传入实体无Execl映射信息";
                return null;
            }

            foreach (DataRow row in dt.Rows)
            {
                T t = new T();
                foreach (PropertyInfo Prty in PropertyInfos)
                {


                    #region 获取指定字段ExeclMappingAttribute
                    string ExeclMappingPropertyKey = $"{TName}_{Prty.Name}_ExeclMappingPropertyKey";
                    ExeclMappingAttribute PrtyExeclMapping = CacheHelper.GetCache<ExeclMappingAttribute>(ExeclMappingPropertyTimeKey);
                    if (PropertyInfos == null)
                    {
                         PrtyExeclMapping = Prty.GetCustomAttribute<ExeclMappingAttribute>(true);
                        CacheHelper.SetCache(ExeclMappingPropertyKey, PropertyInfos);
                    }
                    #endregion

                    if (dt.Columns.Contains(PrtyExeclMapping.ExeclHeadName))
                    {
                        // 判断此属性是否有Setter
                        if (!Prty.CanWrite) continue;//该属性不可写，直接跳出
                        //取值  
                        object value = row[PrtyExeclMapping.ExeclHeadName];
                        //如果非空，则赋给对象的属性  
                        if (value != DBNull.Value && value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            string ConvertValuePropertyKey = $"{TName}_{Prty.Name}_ConvertValuePropertyKey";

                            ConvertValueAttribute appointValueAttribute = CacheHelper.GetCache<ConvertValueAttribute>(ConvertValuePropertyKey);
                            if (appointValueAttribute == null)
                            {
                                appointValueAttribute = Prty.GetCustomAttribute<ConvertValueAttribute>(true);
                                CacheHelper.SetCache(ConvertValuePropertyKey, appointValueAttribute);
                            }

                            if (appointValueAttribute != null)
                            {
                                value = appointValueAttribute.GetConvertValueValue(value);
                            }
                            object objValue = GetValue(Prty, value);
                            Prty.SetValue(t, objValue, null);
                        }

                    }
                }
                listEntity.Add(t);
            }
            return listEntity;
        }

        private string GetTName<T>() where T : IExeclImport, new()
        {
            return typeof(T).Name;
        }

        private object GetValue(PropertyInfo prty, object value)
        {

            if (prty.PropertyType == typeof(int)
                || prty.PropertyType == typeof(int?))
            {
                return System.Convert.ToInt32(value);
            }
            else if (prty.PropertyType == typeof(DateTime)
                || prty.PropertyType == typeof(DateTime?))
            {
                return System.Convert.ToDateTime(value);
            }
            else if (prty.PropertyType == typeof(string))
            {
                return System.Convert.ToString(value);
            }
            else if (prty.PropertyType == typeof(bool))
            {
                return System.Convert.ToBoolean(value);
            }
            return value;
        }
    }
}

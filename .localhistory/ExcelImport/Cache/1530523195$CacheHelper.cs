using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;

namespace ExcelImport.Cache
{
    /// <summary>
    /// 缓存操作类
    /// </summary>
    public class CacheHelper
    {
        /// <summary>
        /// 获取当前应用程序指定CacheKey的Cache值
        /// </summary>
        /// <param name="CacheKey"></param>
        /// <returns></returns>
        public static object GetCache(string CacheKey)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            return objCache[CacheKey];
        }
        public static T GetCache<T>(string CacheKey)
        {
            object objValue = CacheHelper.GetCache(CacheKey);
            if (objValue!=null)
            {

            }
            return default(T);
        }
        /// <summary>
        /// 设置当前应用程序指定CacheKey的Cache值
        /// </summary>
        /// <param name="CacheKey"></param>
        /// <param name="objObject"></param>
        public static void SetCache(string CacheKey, object objObject)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            objCache.Insert(CacheKey, objObject, null, System.Web.Caching.Cache.NoAbsoluteExpiration,
                TimeSpan.FromHours(2));
        }

        public static object RemoveCache(string CacheKey)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            return objCache.Remove(CacheKey);
        }
       

        /// <summary>
        /// 移除指定条件的key值
        /// </summary>
        /// <param name="CacheKey"></param>
        /// <returns></returns>
        public static void RemoveCacheByWhere(string CacheKey)
        {
            System.Web.Caching.Cache objCache = HttpRuntime.Cache;
            IDictionaryEnumerator CacheEnum = objCache.GetEnumerator();
            while (CacheEnum.MoveNext())
            {
                if (CacheEnum.Key.ToString().StartsWith(CacheKey))
                {
                    objCache.Remove(CacheEnum.Key.ToString());
                }
            }
        }

        internal static List<T> GetCacheList<T>(string CacheKey)
        {
            var ObjList = CacheHelper.GetCache(CacheKey);
            if (ObjList != null)
                return ObjList as List<T>;
            else
                return new List<T>();
        }
    }
}

using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelImportDemo
{
    /// <summary>
    /// 文件简单辅助类
    /// 包括旧有日志文件记录、zip压缩等
    /// </summary>
    public static class fileHelper
    {
        /// <summary>
        /// 获取指定时间范围内的日志文件清单
        /// </summary>
        /// <param name="start">起始时间</param>
        /// <param name="end">结束时间</param>
        /// <param name="Type">日志类型编号</param>
        /// <returns>日志文件清单</returns>
        public static List<string> GetLogFile(DateTime start,DateTime end,int Type){
            int days = (end - start).Days-1;//获取时间差
            days = days < 1 ? 0 : days;
            List<string> strarray = new List<string>();

            string Folder = string.Format(@"\logs\{0}\", getLogFolder(Type));//存放根目录
            string PathFolder = System.Web.HttpContext.Current.Server.MapPath(Folder);//硬盘存放根目录

            string PathChild = string.Empty,//相对路径
                   pathSer = string.Empty;//服务器存储路径
                
            for (int i = 0; i <= days; i++) {

                PathChild = string.Format(@"{0}\{1}\{2}", start.AddDays(i).Year, start.AddDays(i).Month, start.AddDays(i).Day);//存放子目录
                pathSer = Path.Combine(PathFolder, PathChild);

                if (!Directory.Exists(pathSer))
                    continue;

                var files = Directory.GetFiles(pathSer);
                foreach (var file in files)
                    strarray.Add(file.Replace(PathFolder, Folder));
            }

            return strarray;
        }
        private static string getLogFolder(int? type = 0)
        {
            string[] group = { "event", "log", "error" };//"操作日志", "登录日志","错误日志"
            type = type < 0 ? 0 : type < group.Length ? type : 0;
            return group[(int)type];
        }
        #region zip
        /// <summary>
        /// 将指定文件夹及其子文件夹内容打包成zip
        /// </summary>
        /// <param name="folderName">需要添加的文件夹：相对路径,开头结尾需要"\"；如果是绝对路径，需要参数isAbsolutFolder=true</param>
        /// <param name="isAbsolutFolder">是否绝对路径</param>
        /// <returns>下载地址</returns>
        public static string PackageFolder(string folderName, bool isAbsolutFolder = false)
        {

            if (string.IsNullOrWhiteSpace(folderName))
                return string.Empty;
            folderName = folderName.Replace("\\", "/").Trim('/'); 
            string NodeBase = folderName;
            if (!isAbsolutFolder)
            {
                NodeBase = System.Web.HttpContext.Current.Server.MapPath(@"\").Replace("\\", "/").TrimEnd('/');
                folderName = Path.Combine(NodeBase, folderName);// NodeBase + folderName;
            }
            folderName = folderName.TrimEnd('/');


            if (!Directory.Exists(folderName))
            {
                return string.Empty;
            }


            DateTime SaveDate = DateTime.Now;

            string SaveFolder = Path.Combine("ZIP", SaveDate.Year.ToString(), SaveDate.Month.ToString(), SaveDate.Day.ToString());
            string FullFolder = Path.Combine(NodeBase, SaveFolder);
            string FileName = Guid.NewGuid().ToString() + ".zip";
            if (!Directory.Exists(FullFolder))
            {
                Directory.CreateDirectory(FullFolder);
            }
            string SavePath = Path.Combine(FullFolder, FileName);
            string Url = "/" + Path.Combine(SaveFolder, FileName);
            try
            {
                var fileList = Directory.EnumerateFiles(folderName, "*", SearchOption.AllDirectories);
                using (ZipFile zip = new ZipFile(SavePath, Encoding.UTF8))
                {
                    zip.AddFiles(fileList, "");
                    zip.Save();
                }
                //using (Package package = Package.Open(SavePath, FileMode.Create))
                //{
                //    foreach (string fileName in fileList)
                //    {

                //        //The path in the package is all of the subfolders after folderName
                //        string pathInPackage;
                //        pathInPackage = Path.GetDirectoryName(fileName).Replace(folderName, string.Empty) + "/" + Path.GetFileName(fileName);

                //        Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri(pathInPackage, UriKind.Relative));
                //        PackagePart packagePartDocument = package.CreatePart(partUriDocument, "", CompressionOption.Maximum);
                //        using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                //        {
                //            fileStream.CopyTo(packagePartDocument.GetStream());
                //        }

                //    }


                //}

                return Url;
            }
            catch //(Exception e)
            {
                return string.Empty;
                //throw new Exception("Error zipping folder " + folderName, e);
            }


        }
        #endregion
        #region 扩展
        /// <summary>
        /// 根据文件相对路径地址，验证并获取物理路径地址
        /// </summary>
        /// <param name="url">需要验证获取的相对地址</param>
        /// <param name="ExtName">扩展文件名称</param>
        /// <param name="ext">需要验证的扩展格式,不设置时不验证直接获取</param>
        /// <returns>返回获取到的物理路径地址，验证失败或不存在时返回string.Empty</returns>
        public static string GetFileDiskUrl(string url, out string ExtName, string[] ext = null)
        {
            ExtName = "";
            if (string.IsNullOrWhiteSpace(url))
                return string.Empty;

            url = System.Web.HttpContext.Current.Server.MapPath(url);
            if (ext != null && ext.Length > 0)
            {
                if (string.IsNullOrWhiteSpace(url) || url.IndexOf('.') < 1)
                    return string.Empty;

                string ExtUrl = Path.GetExtension(url).ToLower();
                ExtName = ExtUrl;
                if (string.IsNullOrWhiteSpace(ExtUrl) || !((System.Collections.IList)ext).Contains(ExtUrl))
                    return string.Empty;
            }

            if (!File.Exists(url))
            {
                return string.Empty;
            }

            return url;
        }
        #endregion
    }
}

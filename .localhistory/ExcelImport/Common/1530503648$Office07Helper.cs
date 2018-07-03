using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;

namespace ExcelImportDemo
{
    /// <summary>
    /// 需要NOPI插件支持
    /// </summary>
    public class Office07Helper
    {
        /// <summary>
        /// 网站根目录所在物理路径(不包括结尾"\")
        /// </summary>
        public static string NodeBase = System.Web.HttpContext.Current.Server.MapPath(@"\").TrimEnd('\\');
        #region 写Excel
        private NPOI.SS.UserModel.ICell SetBorder(NPOI.SS.UserModel.ICell cell)
        {
            cell.CellStyle.BorderLeft = BorderStyle.Thin;
            cell.CellStyle.BorderRight = BorderStyle.Thin;
            cell.CellStyle.BorderTop = BorderStyle.Thin;
            cell.CellStyle.BorderBottom = BorderStyle.Thin;
            cell.CellStyle.Alignment = HorizontalAlignment.Center;
            cell.CellStyle.VerticalAlignment = VerticalAlignment.Center;
            return cell;
        }
        /// <summary>
        /// 根据DataTable创建一个Excel的WorkBook并返回
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <returns>XSSFWorkbook对象</returns>
        public XSSFWorkbook CreateExcelByDataTable(DataTable dt)
        {
            //创建一个工作博
            XSSFWorkbook wk = new XSSFWorkbook();
            //创建以个Sheet
            ISheet tb = wk.CreateSheet("Sheet1");
            
            //创建表头(在第0行)
            IRow row = tb.CreateRow(0);
            #region 根据Datable表头创建Excel表头
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                //创建单元格
                NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                //cell.CellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.DarkRed.Index;
                //cell.CellStyle.FillForegroundColor= NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                cell.SetCellValue(dt.Columns[i].ColumnName);
                tb.AutoSizeColumn(i);//自动调整宽度，貌似对中文支持不好
                SetBorder(cell);
            }
            #endregion
            #region 根据DataTable内容创建Excel内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rows = tb.CreateRow(i+1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = rows.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                    SetBorder(cell);
                }
            }
            #endregion
            return wk;
        }

        /// <summary>
        /// 根据DataTable创建一个Excel的WorkBook并返回
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="columnName">表头</param>
        /// <param name="SheetName">表名称</param>
        /// <returns>XSSFWorkbook对象</returns>
        public XSSFWorkbook CreateExcelByDataTable(DataTable dt, string[] columnName,string SheetName="Sheet1")
        {
            //创建一个工作博
            XSSFWorkbook wk = new XSSFWorkbook();
            //创建以个Sheet
            ISheet tb = wk.CreateSheet(SheetName);
            //创建表头(在第0行)
            IRow row = tb.CreateRow(0);
            #region 表头根据参数
            for (int i = 0; i < columnName.Length; i++)
            {
                //创建单元格
                NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                cell.SetCellValue(columnName[i]);
                cell.CellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.DarkBlue.Index; 
                SetBorder(cell);
            }
            #endregion
            #region 根据DataTable内容创建Excel内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rows = tb.CreateRow(i);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j >= columnName.Length)
                        break;
                    NPOI.SS.UserModel.ICell cell = rows.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                    SetBorder(cell);
                }
            }
            #endregion
            return wk;
        }

        /// <summary>
        /// 根据List(list(string))"创建一个Excel的WorkBook并返回
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="columnName">表头</param>
        /// <param name="columnWidth">列宽度*256</param>
        /// <param name="SheetName">表名称</param>
        /// <returns>XSSFWorkbook对象</returns>
        public XSSFWorkbook CreateExcelByList(System.Collections.Generic.List<System.Collections.Generic.List<object>> dt, string[] columnName,int[] columnWidth, string SheetName = "Sheet1")
        {
            //创建一个工作博
            XSSFWorkbook wk = new XSSFWorkbook();
            //创建以个Sheet
            ISheet tb = wk.CreateSheet(SheetName);
            //设置单元格宽度
            for (int i = 0; i < columnWidth.Length; i++) {
                tb.SetColumnWidth(i,columnWidth[i]*256);
            }
            //创建表头(在第0行)
            IRow row = tb.CreateRow(0);
            #region 表头根据参数
            for (int i = 0; i < columnName.Length; i++)
            {
                //创建单元格
                NPOI.SS.UserModel.ICell cell = row.CreateCell(i);
                cell.SetCellValue(columnName[i]);
                //背景
                //cell.CellStyle.FillForegroundColor = NPOI.XSSF.Util.XSSFColor.DarkTeal.Index;
                //cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                //边框颜色
                //cell.CellStyle.TopBorderColor = NPOI.XSSF.Util.XSSFColor.OliveGreen.Index;
                //cell.CellStyle.LeftBorderColor = NPOI.XSSF.Util.XSSFColor.OliveGreen.Index;
                //cell.CellStyle.BottomBorderColor = NPOI.XSSF.Util.XSSFColor.OliveGreen.Index;
                //cell.CellStyle.RightBorderColor = NPOI.XSSF.Util.XSSFColor.OliveGreen.Index;
                
                cell.CellStyle.IsLocked = true;
                SetBorder(cell);
            }
            #endregion
            #region 根据List<List<object>>内容创建Excel内容
            for (int i = 0; i < dt.Count; i++)
            {
                IRow rows = tb.CreateRow(i+1);
                for (int j = 0; j < dt[i].Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = rows.CreateCell(j);
                    cell.SetCellValue(dt[i][j].ToString());
                    SetBorder(cell);
                }
            }
            #endregion
            return wk;
        }
        /// <summary>
        /// 根据List(list(ExcelOutPut))"创建一个Excel的WorkBook并返回
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="SheetName">表名称</param>
        /// <returns>XSSFWorkbook对象</returns>
        public XSSFWorkbook CreateExcelByList(System.Collections.Generic.List<System.Collections.Generic.List<ExcelOutPut>> dt, string SheetName = "Sheet1")
        {
            //string[] columnName, int[] columnWidth
            System.Collections.Generic.List<System.Collections.Generic.List<object>> dt2 = new System.Collections.Generic.List<System.Collections.Generic.List<object>>();
            System.Collections.Generic.List<object> tempCell;
            System.Collections.Generic.List<string> tempColName=new System.Collections.Generic.List<string>();
            System.Collections.Generic.List<int> tempWidth=new System.Collections.Generic.List<int>();
            //获取第一行标题信息
            for (int j = 0; j < dt[0].Count; j++)
            {
                tempColName.Add(dt[0][j].Title);
                tempWidth.Add(dt[0][j].Len);
            }
            //读取内容
            for (int i = 0; i < dt.Count; i++) {
                tempCell = new System.Collections.Generic.List<object>();
                for (int j = 0; j < dt[i].Count; j++) {
                    tempCell.Add(dt[i][j].Text);
                    //tempColName.Add(dt[i][j].Title);
                    //tempWidth.Add(dt[i][j].Len);
                }
                dt2.Add(tempCell);
            }
            return CreateExcelByList(dt2, tempColName.ToArray(), tempWidth.ToArray(),SheetName);
        }
        /// <summary>
        /// 读取模板，根据List(list(string))"创建一个Excel的WorkBook并返回
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="columnName">表头</param>
        /// <param name="columnWidth">列宽度*256</param>
        /// <param name="ModUrl">模板地址</param>
        /// <param name="startRowIndex">起始行,首行=0，默认0</param>
        /// <param name="RemoveOtherRows">是否移除起始行后的数据，默认=是</param>
        /// <param name="SheetName">表名称</param>
        /// <returns>XSSFWorkbook对象</returns>
        public XSSFWorkbook CreateExcelByList_Mod(System.Collections.Generic.List<System.Collections.Generic.List<object>> dt, string[] columnName, int[] columnWidth, string ModUrl, int startRowIndex = 0,bool RemoveOtherRows=true, string SheetName = "Sheet1")
        {
            startRowIndex=startRowIndex>=0?startRowIndex:0;

            string ExcelModePath = NodeBase + ModUrl;
            if (!File.Exists(ExcelModePath))
            {
                return null;
            }

            //创建一个工作博
            XSSFWorkbook wk = new XSSFWorkbook();

            using (FileStream fileExcelMod = new FileStream(ExcelModePath, FileMode.Open, FileAccess.Read))
            {
                wk = new XSSFWorkbook(fileExcelMod);
                fileExcelMod.Close();
            }
            if (wk == null)
                return null;

            //创建一个Sheet
            ISheet tb = wk.GetSheetAt(0);
            #region 移除起始行后的所有行
            if (RemoveOtherRows&&tb.LastRowNum > startRowIndex + 1)
            {
                for (int rmr = tb.LastRowNum - 1; rmr > startRowIndex; rmr--)
                {
                    tb.ShiftRows(rmr, rmr+1,-1);
                }
            }
            #endregion

            //设置单元格宽度

            //创建表头(在第0行)

            #region 根据List<List<object>>内容创建Excel内容
            for (int i = 0; i < dt.Count; i++)
            {
                IRow rows = tb.CreateRow(i + startRowIndex);
                for (int j = 0; j < dt[i].Count; j++)
                {
                    NPOI.SS.UserModel.ICell cell = rows.CreateCell(j);
                    cell.SetCellValue(dt[i][j].ToString());
                    SetBorder(cell);
                }
            }
            #endregion
            return wk;
        }
        /// <summary>
        /// 保存Excel并返回保存地址
        /// </summary>
        /// <param name="wk">XSSFWorkbook对象</param>
        /// <returns>保存地址</returns>
        public string SaveExcel(XSSFWorkbook wk)
        {

            string Nodes=string.Format(@"\Excel\{0}\{1}\{2}\",DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day);
            string filename =  Guid.NewGuid()+".xlsx";

            if (!Directory.Exists(NodeBase+Nodes))
                Directory.CreateDirectory(NodeBase+Nodes);

            using (FileStream fs = File.OpenWrite(NodeBase+Nodes+filename))
            {
                wk.Write(fs);
            }
            return Nodes+filename;
        }
        /// <summary>
        /// 根据地址，通过DataTable保存excel
        /// </summary>
        /// <param name="wk">XSSFWorkbook对象</param>
        /// <param name="Folder">从网站根目录开始的相对文件夹路径</param>
        /// <returns>返回相对存储地址</returns>
        public string SaveExcelWithFolder(XSSFWorkbook wk,string Folder)
        {
            Folder = Folder.Replace("\\","/");
            string folderFULLAddress = Path.Combine(NodeBase, Folder.Trim('/'));
            if (!Directory.Exists(folderFULLAddress))
                Directory.CreateDirectory(folderFULLAddress);
            string filename = folderFULLAddress.Substring(folderFULLAddress.LastIndexOf('/')+1)+ ".xlsx";
            using (FileStream fs = File.OpenWrite(Path.Combine(folderFULLAddress, filename)))
            {
                wk.Write(fs);
            }
            return Path.Combine(Folder, filename);
        }

        #endregion
        #region 读Excel
        /// <summary>
        /// 读取标准Excel格式
        /// </summary>
        /// <param name="filepath">文件地址</param>
        /// <param name="sheetName">工作博名称</param>
        /// <param name="ReaderRow">从第几行开始</param>
        /// <param name="haveNumber">是否需要计数字段</param>
        /// <returns>DataTable数据集合</returns>
        public DataTable ReaderExcel(string filepath, string sheetName, int ReaderRow = 0, bool haveNumber = false)
        {
            
            DataTable dt = new DataTable();
            using (FileStream fs = File.OpenRead(filepath))
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);
                #region 创建表头
                if (haveNumber)
                {
                    dt.Columns.Add(new DataColumn("DtNumber"));
                }
                //如果ReaderRow大于0,则不是从一行开始读取.创建默认表头
                ISheet sheet = string.IsNullOrWhiteSpace(sheetName) ? wk.GetSheetAt(0) : wk.GetSheet(sheetName);
                if (ReaderRow > 0)
                {
                    
                    for (int i = 0; i < sheet.GetRow(0).LastCellNum; i++)
                    {
                        dt.Columns.Add(new DataColumn("Column" + (i + 1)));
                    }
                }
                else
                {
                    
                    for (int i = 0; i < sheet.GetRow(0).LastCellNum; i++)
                    {
                        dt.Columns.Add(new DataColumn(sheet.GetRow(0).GetCell(i).ToString()));
                    }
                }
                #endregion
                #region 读取入内容

                if (haveNumber)
                {
                    

                    for (int i = ReaderRow; i < sheet.LastRowNum; i++)
                    {
                        DataRow row = dt.NewRow();
                        row[0] = (i + 1);
                        for (int j = 0; j < sheet.GetRow(0).LastCellNum; j++)
                        {
                            row[(j + 1)] = sheet.GetRow(i).GetCell(j).ToString();
                        }
                    }
                }
                else
                {
                    
                    for (int i = ReaderRow; i <= sheet.LastRowNum; i++)
                    {
                        DataRow row = dt.NewRow();
                        for (int j = 0; j < sheet.GetRow(0).LastCellNum; j++)
                        {
                            row[j] = sheet.GetRow(i).GetCell(j).ToString();
                        }
                        dt.Rows.Add(row);
                    }
                }
                #endregion
            }
            return dt;
        }
        #endregion
        #region 读word
        /// <summary>
        /// 读取Word以字符串方式返回
        /// </summary>
        /// <param name="filepath">文档地址</param>
        /// <returns>字符串形式的文档内容</returns>
        public string ReaderWord(string filepath)
        {
            using (FileStream fs = File.OpenRead(filepath))
            {
                XWPFDocument doc = new XWPFDocument(fs);
                XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
                return extractor.ToString();
            }
        }
        /// <summary>
        /// 文档属性
        /// </summary>
        /// <param name="filepath">文档地址</param>
        /// <returns>0.创建者,1分类,2标题</returns>
        public Tuple<string, string, string> GetDocProperties(string filepath)
        {
            using (FileStream fs = File.OpenRead(filepath))
            {
                XWPFDocument doc = new XWPFDocument(fs);
                XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
                CoreProperties t = extractor.GetCoreProperties();
                return new Tuple<string, string, string>(t.Creator, t.Category, t.Title);
            }
        }
        #endregion
        #region 分页压缩导出
        /// <summary>
        /// 分页压缩导出
        /// </summary>
        /// <param name="dtresult">数据集</param>
        /// <param name="PartName">导出数据归类（中文无其他符号）</param>
        /// <param name="SaveFolder">存储地址（相对地址，且以/开头，/结束）,没有时将使用自带地址</param>
        /// <returns>导出地址，错误或无数据时返回null</returns>
        public string OutputExcelWithPageForZip(DataTable dtresult, string PartName = "导出数据", string SaveFolder = "")
        {
            try
            {
                PartName = string.IsNullOrWhiteSpace(PartName) ? "导出数据" : PartName;
                string FilePathFolder = !string.IsNullOrWhiteSpace(SaveFolder) ? SaveFolder : string.Format("/OPExcel/{0}/{1}/", PartName.ToPinYin("OPExcel",true).Replace(" ", ""), DateTime.Now.ToString("yyyyMMdd"));
                FilePathFolder = string.Format("/{0}/{1}/", FilePathFolder.Trim('/'), Guid.NewGuid().ToString());

                string FilePath = "";
                System.Collections.Generic.List<string> FilePathArray = new System.Collections.Generic.List<string>();


                var rows = dtresult.Rows.Cast<DataRow>();
                DataTable pageDT = dtresult.Clone();

                int PageSize = 3000;
                int totalPage = (dtresult.Rows.Count / PageSize) + (dtresult.Rows.Count % PageSize > 0 ? 1 : 0);//总页码

                for (int PageIndex = 1; PageIndex <= totalPage; PageIndex++)
                {
                    var curRows = rows.Skip((PageIndex - 1) * PageSize).Take(PageSize).ToArray();
                    foreach (var r in curRows)
                        pageDT.ImportRow(r);

                    FilePath = Path.Combine(FilePathFolder, string.Format("{0}_第{1}部分（{2}条数据）", PartName, PageIndex, pageDT.Rows.Count));
                    SaveExcelWithFolder(CreateExcelByDataTable(pageDT), FilePath);
                    pageDT.Clear();
                    FilePathArray.Add(FilePath);
                }
                if (FilePathArray.Count < 1) return null;

                return fileHelper.PackageFolder(FilePathFolder);
            }
            catch(Exception ex) {
                throw ex;
                //return null;
            }
        }
        #endregion
    }
}

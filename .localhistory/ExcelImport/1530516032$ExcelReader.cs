using System.Data;
using System.IO;

namespace ExcelImport
{
    public class ExcelReader
    {

        /// <summary>
        /// 读取标准Excel格式
        /// </summary>
        /// <param name="filepath">文件地址</param>
        /// <param name="sheetName">工作博名称</param>
        /// <param name="ReaderRow">从第几行开始</param>
        /// <param name="haveNumber">是否需要计数字段</param>
        /// <returns>DataTable数据集合</returns>
        protected (string[] head, DataTable dt) ReaderExcel(string filepath, string sheetName)
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



                for (int i = 0; i <= sheet.LastRowNum; i++)
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
    }
}

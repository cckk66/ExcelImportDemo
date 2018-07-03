namespace ExcelImport.Model
{
    /// <summary>
    /// Excel导出格式
    /// </summary>
    public class ExcelOutPut
    {
        /// <summary>
        /// 内容
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 列宽
        /// </summary>
        public int Len { get; set; }
    }
}

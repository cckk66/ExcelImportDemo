using System.Data;

namespace ExcelImport
{
    public abstract class ExcelImportHelper
    {
        public int MaxLeng { get; set; } = 500;
        public string ErrorMsg { get; set; } = "";

        private DataTable ExeclToDataTable(string Url)
        {

        }
    }
}

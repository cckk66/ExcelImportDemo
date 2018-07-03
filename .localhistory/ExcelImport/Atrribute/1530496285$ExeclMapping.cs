namespace ExcelImport.Atrribute
{
    public class ExeclMapping : System.Attribute
    {
        public string ExeclHeadName { get; set; }
        public ExeclMapping(string _ExeclHeadName)
        {
            ExeclHeadName = _ExeclHeadName;
        }
    }
}

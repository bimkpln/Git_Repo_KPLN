namespace KPLN_Library_DataBase.Controll
{
    public class DbSubDepartmentInfo
    {
        public DbSubDepartmentInfo(int id, string name, string code, string codeUS)
        {
            Id = id;
            Name = name;
            Code = code;
            CodeUS = codeUS;
        }
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
        public string CodeUS { get; }
    }
}

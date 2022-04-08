namespace KPLN_DataBase.Controll
{
    public class DbDepartmentInfo
    {
        public DbDepartmentInfo(int id, string name, string code)
        {
            Id = id;
            Name = name;
            Code = code;
        }
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
    }
}

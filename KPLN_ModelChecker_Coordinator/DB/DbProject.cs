namespace KPLN_ModelChecker_Coordinator.DB
{
    public class DbProject
    {
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
        public DbProject(int id, string name, string code)
        {
            Id = id;
            Name = name;
            Code = code;
        }
    }
}

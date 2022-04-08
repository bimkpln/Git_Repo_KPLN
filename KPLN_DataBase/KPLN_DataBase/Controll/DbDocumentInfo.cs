using KPLN_DataBase.Collections;
using System.Collections.ObjectModel;

namespace KPLN_DataBase.Controll
{
    public class DbDocumentInfo
    {
        public DbDocumentInfo(int id, string path, string name, string code, int project, int department, ObservableCollection<DbProject> projects, ObservableCollection<DbSubDepartment> departments)
        {
            Id = id;
            Path = path;
            Name = name;
            foreach (DbProject p in projects)
            {
                if (p.Id == project)
                { Project = p; }
            }
            foreach (DbSubDepartment d in departments)
            {
                if (d.Id == department)
                { Department = d; }
            }
            Code = code;
        }
        public int Id { get; }
        public string Path { get; }
        public string Name { get; }
        public DbProject Project { get; }
        public DbSubDepartment Department { get; }
        public string Code { get; }
    }
}

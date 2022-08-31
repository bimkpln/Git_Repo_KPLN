using KPLN_Library_DataBase.Collections;
using KPLN_Library_DataBase.Controll;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Library_DataBase
{
    public static class DbControll
    {
        public static ObservableCollection<DbDepartment> Departments { get; private set; }
        public static ObservableCollection<DbDocument> Documents { get; private set; }
        public static ObservableCollection<DbProject> Projects { get; private set; }
        public static ObservableCollection<DbSubDepartment> SubDepartments { get; private set; }
        public static ObservableCollection<DbUser> Users { get; private set; }
        public static DbUser CurrentUser { get; private set; }
        public static void Update()
        {
            _Departments = SQLiteDBUtills.GetDepartmentInfo();
            Departments = DbDepartment.GetAllDepartments(_Departments);
            _SubDepartments = SQLiteDBUtills.GetSubDepartmentInfo();
            SubDepartments = DbSubDepartment.GetAllSubDepartments(_SubDepartments);
            _Users = SQLiteDBUtills.GetUserInfo(Departments);
            Users = DbUser.GetAllUsers(_Users);
            _Projects = SQLiteDBUtills.GetProjectInfo(Users);
            Projects = DbProject.GetAllProjects(_Projects);
            _Documents = SQLiteDBUtills.GetDocumentInfo(SubDepartments, Projects);
            Documents = DbDocument.GetAllDocuments(_Documents);
            string currentUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();
            foreach (DbUser user in Users)
            {
                if (user.SystemName == currentUserName)
                {
                    CurrentUser = user;
                }
            }
            foreach (DbProject project in Projects)
            {
                project.JoinDocumentsFromList(Documents);
            }
        }
        public static ObservableCollection<DbDepartmentInfo> _Departments { get; set; }
        public static ObservableCollection<DbDocumentInfo> _Documents { get; set; }
        public static ObservableCollection<DbProjectInfo> _Projects { get; set; }
        public static ObservableCollection<DbSubDepartmentInfo> _SubDepartments { get; set; }
        public static ObservableCollection<DbUserInfo> _Users { get; set; }
        public static ObservableCollection<DbDocument> GetDocumentsByProject(DbProject project)
        {
            ObservableCollection<DbDocument> documents = new ObservableCollection<DbDocument>();
            foreach (DbDocument doc in Documents)
            {
                if (doc.Project.Id == project.Id)
                {
                    documents.Add(doc);
                }
            }
            return documents;
        }
        public static ObservableCollection<DbUser> GetUsersByNames(List<string> userData)
        {
            ObservableCollection<DbUser> users = new ObservableCollection<DbUser>();
            foreach (DbUser user in Users)
            {
                foreach (string userName in userData)
                {
                    if (user.SystemName == userName)
                    {
                        users.Add(user);
                    }
                }
            }
            return users;
        }
    }
}

using KPLN_Library_DataBase.Collections;
using KPLN_Library_DataBase.Controll;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.Linq;

namespace KPLN_Library_DataBase
{
    public static class DbControll
    {
        /// <summary>
        /// Путь к основной базе данных. Используется во всех загружаемых модулях для централизованного управления БД
        /// </summary>
        public static readonly string MainDBPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db";

        /// <summary>
        /// Подключение к основной базе данных. Используется во всех загружаемых модулях для централизованного управления БД
        /// </summary>
        public static readonly string MainDBConnection = string.Format(@"Data Source=" + MainDBPath + ";Version=3;");

        public static ObservableCollection<DbDepartment> Departments { get; private set; }
        
        internal static ObservableCollection<DbDepartmentInfo> DepartmentInfos { get; set; }
        
        public static ObservableCollection<DbDocument> Documents { get; private set; }
        
        internal static ObservableCollection<DbDocumentInfo> DocumentInfos { get; set; }
        
        public static ObservableCollection<DbProject> Projects { get; set; }

        internal static ObservableCollection<DbProjectInfo> ProjectInfos { get; set; }
        
        public static ObservableCollection<DbSubDepartment> SubDepartments { get; private set; }

        internal static ObservableCollection<DbSubDepartmentInfo> SubDepartmentInfos { get; set; }
        
        public static ObservableCollection<DbUser> Users { get; private set; }
        
        internal static ObservableCollection<DbUserInfo> UserInfos { get; set; }
        
        public static DbUser CurrentUser { get; private set; }
        
        /// <summary>
        /// Обновление данных из базы данных КПЛН
        /// </summary>
        public static void Update()
        {
            DepartmentInfos = SQLiteDBUtills.GetDepartmentInfo();
            Departments = DbDepartment.GetAllDepartments(DepartmentInfos);
            
            UserInfos = SQLiteDBUtills.GetUserInfo(Departments);
            Users = DbUser.GetAllUsers(UserInfos);

            ProjectInfos = SQLiteDBUtills.GetProjectInfo(Users);
            Projects = DbProject.GetAllProjects(ProjectInfos);

            SubDepartmentInfos = SQLiteDBUtills.GetSubDepartmentInfo();
            SubDepartments = DbSubDepartment.GetAllSubDepartments(SubDepartmentInfos);

            DocumentInfos = SQLiteDBUtills.GetDocumentInfo(SubDepartments, Projects);
            Documents = DbDocument.GetAllDocuments(DocumentInfos);
            
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

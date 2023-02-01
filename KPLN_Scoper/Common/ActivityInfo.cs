using Autodesk.Revit.DB;
using KPLN_Library_DataBase.Collections;
using System;
using System.IO;
using static KPLN_Scoper.Common.Collections;
using static KPLN_Loader.Output.Output;

namespace KPLN_Scoper.Common
{
    public class ActivityInfo
    {
        public string DocumentTitle { get; set; } = string.Empty;
        
        public int ProjectId { get; set; }
        
        public int DocumentId { get; set; }
        
        public double Value { get; set; } = 1.0;
        
        public BuiltInActivity Type { get; set; }
        
        public string Time { get; set; } = DateTime.Now.ToString();
        
        public ActivityInfo(Document doc, BuiltInActivity type)
        {
            if (doc.IsDetached || !doc.IsWorkshared) 
            { 
                throw new Exception("Внимание: Документ не для совместной работы!"); 
            }
            
            Type = type;
            
            string filename = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())).FullName;
            
            View actView = doc.ActiveView;
            
            if (actView != null)
            {
                DocumentTitle = string.Format("{0} : {1}", filename, actView.Name);
                SetInfoDocValues(filename);
            }
            else
            {
                DocumentTitle = string.Format("{0}", filename);
                SetInfoDocValues(filename);
            }
        }
        
        public ActivityInfo(int doc, int project, BuiltInActivity type, string title, int skip_amount)
        {
            DocumentTitle = title;
            Type = type;
            DocumentId = doc;
            ProjectId = project;
            Value = 1.0;
            if (type == BuiltInActivity.ActiveDocument)
            {
                for (int i = 0; i < skip_amount; i++)
                {
                    if (type == BuiltInActivity.ActiveDocument)
                    { Value = Value * 0.9; }
                }
            }
            if (type == BuiltInActivity.DocumentChanged)
            {
                for (int i = 0; i < skip_amount - 1; i++)
                {
                    if (type == BuiltInActivity.DocumentChanged)
                    { Value = Value * 1.1; }
                }
            }
        }

        private void SetInfoDocValues(string filename)
        {
            DocumentId = -1;
            ProjectId = -1;
            foreach (DbDocument docu in KPLN_Library_DataBase.DbControll.Documents)
            {
                if (docu.Path == String.Empty)
                    throw new Exception($"Разработчик скоро это устранит, сообщение адресовано ему:" +
                        $"\n У элемента с id {docu.Id} - проблемы с определением пути. Проверь заполнение БД!");
                
                if (new FileInfo(docu.Path).FullName == filename)
                {
                    DocumentId = docu?.Id ?? -1;
                    ProjectId = docu.Project?.Id ?? -1;
                }
            }
        }
    }
}

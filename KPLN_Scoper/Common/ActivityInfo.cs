using Autodesk.Revit.DB;
using KPLN_Library_DataBase.Collections;
using System;
using System.IO;
using static KPLN_Scoper.Common.Collections;

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
            Type = type;
            if (doc.IsDetached || !doc.IsWorkshared) { throw new Exception("Документ не подходит для совместной работы!"); }
            string filename = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())).FullName;
            DocumentTitle = string.Format("{0} : {1}", filename, doc.ActiveView.Name);
            DocumentId = -1;
            ProjectId = -1;
            try
            {
                foreach (DbDocument docu in KPLN_Library_DataBase.DbControll.Documents)
                {
                    try
                    {
                        if (new FileInfo(docu.Path).FullName == filename)
                        {
                            DocumentId = docu.Id;
                            if (docu.Project != null)
                            {
                                ProjectId = docu.Project.Id;
                            }
                        }
                    }
                    catch (Exception)
                    { }
                }
            }
            catch (Exception)
            { }

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
    }
}

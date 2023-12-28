using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;

namespace KPLN_Library_Forms.Common
{
    public class ElementEntity
    {
        public ElementEntity(object elem)
        {
            Element = elem;

            // Анбоксинг на Element Revit
            if (elem is Element element)
            {
                Name = element.Name;
            }

            // Анбоксинг на DbProject KPLN_Library_DataBase
            if (elem is DBProject dbProject)
            {
                Name = $"{dbProject.Name}. Стадия: {dbProject.Stage}";
                Tooltip = dbProject.MainPath;
            }
        }

        public object Element { get; private set; }

        public string Name { get; private set; }

        public string Tooltip { get; private set; }
    }
}

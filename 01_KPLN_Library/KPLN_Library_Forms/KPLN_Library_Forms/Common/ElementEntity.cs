using Autodesk.Revit.DB;
using KPLN_Library_DataBase.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Library_Forms.Common
{
    public class ElementEntity
    {
        public ElementEntity(object elem)
        {
            Element = elem;

            // Анбоксинг на Element Revit
            if (elem is Element)
            {
                Element element = (Element)elem;
                Name = element.Name;
                return;
            }

            // Анбоксинг на DbProject KPLN_Library_DataBase
            if (elem is DbProject)
            {
                DbProject dbProject = (DbProject)elem;
                Name = dbProject.Name;
                return;
            }
        }
        
        public object Element { get; private set; }
        
        public string Name { get; private set; }
    }
}

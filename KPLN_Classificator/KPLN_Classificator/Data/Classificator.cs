using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace KPLN_Classificator
{
    [Serializable]
    public class Classificator
    {
        public List<string> paramsValues;
        public BuiltInCategory BuiltInName;
        
        /// <summary>
        /// Имя семейства для фильтрации
        /// </summary>
        public string FamilyName;
        
        /// <summary>
        /// Имя типа для фильтрации
        /// </summary>
        public string TypeName;
        
        /// <summary>
        /// Парамтер для фильтрации
        /// </summary>
        public string ParameterName;
        
        /// <summary>
        /// Значение параметра для фильтрации
        /// </summary>
        public string ParameterValue;

        public Classificator()
        {

        }
    }
}

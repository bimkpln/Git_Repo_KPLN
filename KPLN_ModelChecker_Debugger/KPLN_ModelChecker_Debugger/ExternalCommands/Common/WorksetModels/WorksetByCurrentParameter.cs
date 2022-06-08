using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Debugger.ExternalCommands.Common.WorksetModels
{
    public class WorksetByCurrentParameter
    {
        /// <summary>
        /// Имя рабочего набора
        /// </summary>
        public string WorksetName;

        /// <summary>
        /// Список категорий
        /// </summary>
        public List<BuiltInCategory> BuiltInCategories = new List<BuiltInCategory>();

        /// <summary>
        /// Список части имен семейств
        /// </summary>
        public List<string> FamilyNames = new List<string>();

        /// <summary>
        /// Список части имен типов
        /// </summary>
        public List<string> TypeNames = new List<string>();

        /// <summary>
        /// Список параметров и их значений (спец. класс)
        /// </summary>
        public List<SelectedParameter> SelectedParameters = new List<SelectedParameter>();

        /// <summary>
        /// Получение рабочего набора из проекта
        /// </summary>
        public Workset GetWorkset(Document doc)
        {
            IList<Workset> userWorksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets();
            bool isUniqueWSet = WorksetTable.IsWorksetNameUnique(doc, WorksetName);

            if (!isUniqueWSet)
            {
                Workset wSet = new FilteredWorksetCollector(doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .ToWorksets()
                    .Where(w => w.Name == WorksetName)
                    .First();
                return wSet;
            }
            else
            {
                Workset wSet = Workset.Create(doc, WorksetName);
                return wSet;
            }
        }

        /// <summary>
        /// Назначение указанного рабочего набора
        /// </summary>
        public static void SetWorkset(Element elem, Workset w)
        {
            Parameter wSParam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            if (wSParam == null) return;
            if (wSParam.IsReadOnly) return;

            bool elemNonGroup = (elem.GroupId == null) || (elem.GroupId == ElementId.InvalidElementId);
            if (elemNonGroup)
            {
                wSParam.Set(w.Id.IntegerValue);
            }
            else
            {
                Group gr = elem.Document.GetElement(elem.GroupId) as Group;
                SetWorkset(gr, w);
            }
        }
    }
}

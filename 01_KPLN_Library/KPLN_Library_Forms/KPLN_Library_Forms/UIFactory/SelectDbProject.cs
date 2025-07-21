using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Library_Forms.UIFactory
{
    /// <summary>
    /// Выбор проекта из базы данных KPLN
    /// </summary>
    public static class SelectDbProject
    {
        /// <summary>
        /// Запуск окна выбора проекта
        /// </summary>
        /// <returns>Возвращает выбранный проект</returns>
        public static ElementSinglePick CreateForm(int rVersion, bool showClosedProjects = false)
        {
            ObservableCollection<ElementEntity> projects = new ObservableCollection<ElementEntity>();

            IEnumerable<DBProject> projectsColl = null;
            switch (showClosedProjects)
            {
                case true:
                    projectsColl = DBMainService.ProjectDbService.GetDBProjects_ByRVersion(rVersion);
                    break;
                case false:
                    projectsColl = DBMainService.ProjectDbService.GetDBProjects_ByRVersionANDOpened(rVersion);
                    break;
            }

            foreach (DBProject prj in projectsColl)
            {
                if (prj.Name.Equals(null))
                    throw new Exception($"KPLN_Exception: Ошибка в заполнении БД - у элемента проекта с id: {prj.Id} нет имени");
                else if (prj.Code.Equals(null))
                    throw new Exception($"KPLN_Exception: Ошибка в заполнении БД - у элемента проекта с id: {prj.Id} нет имени");
                else if (prj.Code != "BIM")
                    projects.Add(new ElementEntity(prj, prj.MainPath));
            }

            ElementSinglePick pickForm = new ElementSinglePick(projects.OrderBy(p => p.Name), "Выбери проект");

            return pickForm;
        }
    }
}

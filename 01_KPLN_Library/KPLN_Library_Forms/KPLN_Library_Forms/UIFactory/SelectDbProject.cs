using KPLN_Library_DataBase;
using KPLN_Library_DataBase.Collections;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System;
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
        /// <returns>Возвращает выбранный проект, или null, если нужно выбрать всё</returns>
        /// <exception cref="Exception"></exception>
        public static ElementPick CreateForm()
        {
            DbControll.Update();

            ObservableCollection<ElementEntity> projects = new ObservableCollection<ElementEntity>();
            foreach (DbProject prj in DbControll.Projects)
            {
                if (prj.Name.Equals(null))
                {
                    throw new Exception($"KPLN_Exception: Ошибка в заполнении БД - у элемента проекта нет имена");
                }
                else if (prj.Code.Equals(null))
                {
                    throw new Exception($"KPLN_Exception: Ошибка в заполнении БД - у элемента проекта нет имена");
                }
                else if (prj.Code != "BIM")
                {
                    projects.Add(new ElementEntity(prj));
                }
            }

            ElementPick _pickForm = new ElementPick(projects.OrderBy(p => p.Name));

            return _pickForm;
        }
    }
}

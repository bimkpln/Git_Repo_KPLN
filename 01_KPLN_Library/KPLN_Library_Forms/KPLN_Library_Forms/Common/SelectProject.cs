using KPLN_Library_DataBase;
using KPLN_Library_DataBase.Collections;
using KPLN_Library_DataBase.Controll;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_Library_Forms.Common
{
    /// <summary>
    /// Выбор проекта из базы данных
    /// </summary>
    public static class SelectProject
    {

        /// <summary>
        /// Запуск окна выбора проекта
        /// </summary>
        /// <returns>Возвращает выбранный проект, или null, если нужно выбрать всё</returns>
        /// <exception cref="Exception"></exception>
        public static FormSinglePick CreateForm()
        {
            DbControll.Update();
            
            ObservableCollection<DbProject> projects = new ObservableCollection<DbProject>();
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
                    projects.Add(prj);
                }
            }

            FormSinglePick _pickForm = new FormSinglePick(projects.OrderBy(p => p.Name));

            return _pickForm;
        }
    }
}

﻿using System.Collections.Generic;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.WPFItems
{
    /// <summary>
    /// Спец. класс-обертка, для подготовки данных для передачи в отчет
    /// </summary>
    public sealed class WPFReportCreator
    {
        public WPFReportCreator(IEnumerable<WPFEntity> wpfEntityColl, string checkName, string logLastRun)
        {
            WPFEntityCollection = wpfEntityColl;
            int counter = 0;
            foreach (WPFEntity w in WPFEntityCollection)
                w.Header = $"#{++counter} {w.Header}";

            CheckName = checkName;
            LogLastRun = logLastRun;
        }

        public WPFReportCreator(IEnumerable<WPFEntity> wpfEntityColl, string checkName, string logLastRun, string logMarker) : this(wpfEntityColl, checkName, logLastRun)
        {
            LogMarker = logMarker;
        }

        /// <summary>
        /// Коллекция элементов отчета (WPFEntity)
        /// </summary>
        public IEnumerable<WPFEntity> WPFEntityCollection { get; private set; }

        /// <summary>
        /// Коллеция описаний для фильтрации (текстовые значения, по которым группируются элементы)
        /// </summary>
        public HashSet<string> FiltrationCollection { get; private set; } = new HashSet<string>() { "Необработанные предупреждения" };

        /// <summary>
        /// Имя проверки
        /// </summary>
        public string CheckName { get; }

        /// <summary>
        /// Данные лога последнего запуска
        /// </summary>
        public string LogLastRun { get; }

        /// <summary>
        /// Данные лога ключевой пометки (маркера)
        /// </summary>
        public string LogMarker { get; } = null;

        /// <summary>
        /// Обновить коллекцию WPFEntity с указанием параметра для фильтрации по категории Revit
        /// </summary>
        public void SetWPFEntityFiltration_ByCategory()
        {
            foreach (WPFEntity w in WPFEntityCollection)
            {
                string catName = w.CategoryName;
                if (!string.IsNullOrEmpty(catName))
                    w.FiltrationDescription = catName;
                else
                    w.FiltrationDescription = "Ошибка определния категории";

                FiltrationCollection.Add(w.FiltrationDescription);
            }
        }

        /// <summary>
        /// Обновить коллекцию WPFEntity с указанием параметра для фильтрации по статусу ошибки
        /// </summary>
        public void SetWPFEntityFiltration_ByStatus()
        {
            foreach (WPFEntity w in WPFEntityCollection)
            {
                switch (w.CurrentStatus)
                {
                    case Status.AllmostOk:
                        w.FiltrationDescription = "Почти хорошо";
                        break;
                    case Status.LittleWarning:
                        w.FiltrationDescription = "Обрати внимание";
                        break;
                    case Status.Warning:
                        w.FiltrationDescription = "Предупреждение";
                        break;
                    case Status.Error:
                        w.FiltrationDescription = "Ошибка";
                        break;
                    case Status.Approve:
                        w.FiltrationDescription = "Допустимое";
                        break;
                }
                FiltrationCollection.Add(w.FiltrationDescription);
            }
        }

        /// <summary>
        /// Обновить коллекцию WPFEntity с указанием параметра для фильтрации по заголовку в ошибке
        /// </summary>
        public void SetWPFEntityFiltration_ByErrorHeader()
        {
            foreach (WPFEntity w in WPFEntityCollection)
            {
                w.FiltrationDescription = w.ErrorHeader;
                FiltrationCollection.Add(w.FiltrationDescription);
            }
        }

        /// <summary>
        /// Обновить коллекцию WPFEntity с указанием параметра для фильтрации по Id-элемента
        /// </summary>
        public void SetWPFEntityFiltration_ByElementId()
        {
            foreach (WPFEntity w in WPFEntityCollection)
            {
                w.FiltrationDescription = w.ElementId.IntegerValue.ToString();
                FiltrationCollection.Add(w.FiltrationDescription);
            }
        }

        /// <summary>
        /// Обновить коллекцию WPFEntity с указанием параметра для фильтрации по имени элемента
        /// </summary>
        public void SetWPFEntityFiltration_ByElementName()
        {
            foreach (WPFEntity w in WPFEntityCollection)
            {
                w.FiltrationDescription = w.ElementName;
                FiltrationCollection.Add(w.FiltrationDescription);
            }
        }
    }
}

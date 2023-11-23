using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    internal class MonitoringAndPinnerSearcher
    {
        /// <summary>
        /// Проверка осей и уровней
        /// </summary>
        public static IEnumerable<WPFEntity> CheckMainLines(UIApplication uiapp, Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (Element element in elemColl)
            {
                if (element.IsMonitoringLinkElement())
                {
                    RevitLinkInstance link = null;
                    foreach (ElementId i in element.GetMonitoredLinkElementIds())
                    {
                        link = uiapp.ActiveUIDocument.Document.GetElement(i) as RevitLinkInstance;
                        if (link == null)
                        {
                            result.Add(new WPFEntity(
                                element,
                                Status.Error,
                                "Ошибка мониторинга",
                                $"Связь не найдена: «{element.Name}»",
                                true,
                                false,
                                $"Элементу с ID {element.Id} необходимо исправить мониторинг"));
                        }
                        else if (!link.Name.ToLower().Contains("разб"))
                        {
                            result.Add(new WPFEntity(
                                element,
                                Status.Error,
                                "Ошибка мониторинга",
                                $"Мониторинг не из разбивочного файла: «{element.Name}»",
                                true,
                                false,
                                $"Элементу с ID {element.Id} необходимо исправить мониторинг, сейчас он присвоен связи {link.Name}"));
                        }
                    }
                }
                else
                {
                    result.Add(new WPFEntity(
                        element,
                        Status.Error,
                        "Отсутсвие мониторинга",
                        $"Элементу с ID {element.Id} необходимо задать мониторинг",
                        true,
                        false));
                }
                
                if (!element.Pinned)
                {
                    result.Add(new WPFEntity(
                        element,
                        Status.Error,
                        "Отсутсвие прикрепления (PIN)",
                        $"Элемент не прикреплен: «{element.Name}»",
                        true,
                        false,
                        $"Элемент с ID {element.Id} необходимо прикрепить (PIN)"));
                }
            }

            return result;
        }
    }
}

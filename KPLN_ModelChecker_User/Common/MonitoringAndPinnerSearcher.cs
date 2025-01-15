using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.WPFItems;
using System.Collections.Generic;

namespace KPLN_ModelChecker_User.Common
{
    internal class MonitoringAndPinnerSearcher
    {
        /// <summary>
        /// Проверка осей и уровней
        /// </summary>
        public static IEnumerable<WPFEntity> CheckMainLines(ExtensibleStorageEntity esEntity, UIApplication uiapp, Document doc, Element[] elemColl)
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
                                esEntity,
                                element,
                                "Ошибка мониторинга",
                                $"Связь не найдена: «{element.Name}»",
                                $"Элементу с ID {element.Id} необходимо исправить мониторинг",
                                false));
                        }
                        else if (!link.Name.ToLower().Contains("разб")
                            && !link.Name.Contains("СЕТ_1_1-3_00_РФ"))
                        {
                            result.Add(new WPFEntity(
                                esEntity,
                                element,
                                "Ошибка мониторинга",
                                $"Мониторинг не из разбивочного файла: «{element.Name}»",
                                $"Элементу с ID {element.Id} необходимо исправить мониторинг, сейчас он присвоен связи {link.Name}",
                                false));
                        }
                    }
                }
                else
                {
                    result.Add(new WPFEntity(
                        esEntity,
                        element,
                        "Отсутсвие мониторинга",
                        $"Элементу с ID {element.Id} необходимо задать мониторинг",
                        string.Empty,
                        false));
                }

                if (!element.Pinned)
                {
                    result.Add(new WPFEntity(
                        esEntity,
                        element,
                        "Отсутсвие прикрепления (PIN)",
                        $"Элемент не прикреплен: «{element.Name}»",
                        $"Элемент с ID {element.Id} необходимо прикрепить (PIN)",
                        false));
                }
            }

            return result;
        }
    }
}

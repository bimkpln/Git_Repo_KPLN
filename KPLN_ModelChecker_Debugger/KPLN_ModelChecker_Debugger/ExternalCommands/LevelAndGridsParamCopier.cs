using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal sealed class LevelAndGridsParamCopier : IExternalCommand
    {
        private static string _sectParamName;
        private static string _levelParamName;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            #region Настраиваю параметры в зависимости от проекта
            string docPath = doc.Title.ToUpper();
            int userDepartment = Module.CurrentDBUser.SubDepartmentId;

            bool isSET = docPath.StartsWith("СЕТ_1");
            // Посадить на конфиг под каждый файл
            if (userDepartment == 2 || userDepartment == 8 && docPath.Contains("АР"))
            {
                if (docPath.StartsWith("ИЗМЛ"))
                {
                    _levelParamName = "КП_О_Этаж";
                    _sectParamName = "КП_О_Секция";
                }
                if (isSET)
                {
                    _levelParamName = "СМ_Этаж";
                    _sectParamName = "СМ_Секция";
                }
            }
            else if (userDepartment == 3 || userDepartment == 8 && docPath.Contains("КР"))
            {
                if (docPath.StartsWith("ИЗМЛ"))
                {
                    _levelParamName = "О_Этаж";
                    _sectParamName = "КП_О_Секция";
                }
                if (isSET)
                {
                    _levelParamName = "СМ_Этаж";
                    _sectParamName = "СМ_Секция";
                }
            }
            else if (userDepartment == 4
                     || userDepartment == 5
                     || userDepartment == 6
                     || userDepartment == 7
                     || userDepartment == 8
                     && (docPath.Contains("ОВ")
                         || docPath.Contains("ВК")
                         || docPath.Contains("АУПТ")
                         || docPath.Contains("ЭОМ")
                         || docPath.Contains("СС")
                         || docPath.Contains("АК")
                         || docPath.Contains("АВ")))
            {
                if (docPath.StartsWith("ОБДН"))
                {
                    _levelParamName = "КП_О_Этаж";
                    _sectParamName = "КП_О_Секция";
                }
                if (docPath.StartsWith("ИЗМЛ"))
                {
                    _levelParamName = "КП_О_Этаж";
                    _sectParamName = "КП_О_Секция";
                }
                if (isSET)
                {
                    _levelParamName = "СМ_Этаж";
                    _sectParamName = "СМ_Секция";
                }
                if (docPath.StartsWith("ПШМ1"))
                {
                    _levelParamName = "КП_О_Этаж";
                    _sectParamName = "КП_О_Секция";
                }
            }
            else
            {
                TaskDialog td = new TaskDialog("ОШИБКА: Выполни инструкцию")
                {
                    MainContent = "Ошибка определения проекта/пользователя. Обратись в BIM-отдел",
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation
                };
                td.Show();
                return Result.Failed;
            }
            #endregion

            #region Получаю сетки из связей
            List<LinkGridLevelData> linkGridLevelDataElemColl = new List<LinkGridLevelData>();
            RevitLinkInstance[] rvtLinkInstsColl = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                // Слабое место - имена файлов могут отличаться из-за требований Заказчика
                .Where(lm => lm.Name.ToUpper().Contains("_РАЗБ") 
                    // Отлов проектов с требованиями к именованиям разб. Файлам
                    || lm.Name.Contains("СЕТ_1_1-3_00_РФ"))
                .Cast<RevitLinkInstance>()
                .ToArray();
            
            foreach (RevitLinkInstance rvtLinkInst in rvtLinkInstsColl)
            {
                Document linkDoc = rvtLinkInst.GetLinkDocument();
                
                linkGridLevelDataElemColl.AddRange(
                    new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Grids)
                    .WhereElementIsNotElementType()
                    .Select(e => new LinkGridLevelData() { Element = e, RevLinkInstance = rvtLinkInst }));

                // Кастомизация под проект
                if (isSET)
                    linkGridLevelDataElemColl.AddRange(
                        new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Levels)
                        .WhereElementIsNotElementType()
                        .Where(lvl => !lvl.Name.Contains("КР"))
                        .Select(e => new LinkGridLevelData() { Element = e, RevLinkInstance = rvtLinkInst }));
                else
                    linkGridLevelDataElemColl.AddRange(
                        new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Levels)
                        .WhereElementIsNotElementType()
                        .Select(e => new LinkGridLevelData() { Element = e, RevLinkInstance = rvtLinkInst }));
            }

            if (!linkGridLevelDataElemColl.Any())
            {
                Print("Не удалось получить оси и уровни из разбивочного файла. Нужно загрузить файл в проект. Если файлы подгружены, то скинь проблему разработчику",
                    MessageType.Error);
                return Result.Cancelled;
            }
            #endregion

            #region Получаю сетки из проекта
            IEnumerable<Element> prjGridsElemEnum = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();
            Element[] prjGridsElemColl = prjGridsElemEnum
                .Concat(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType())
                .ToArray();
            #endregion

            string userResultSectionsMsg;
            string userResultLevelsMsg;
            HashSet<string> resultGridSections = new HashSet<string>();
            HashSet<string> resultLevels = new HashSet<string>();
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("KPLN: Копировать пар-ры сеток");

                foreach (Element prjGridLevel in prjGridsElemColl)
                {
                    #region Подготовка и проверка элемента
                    IList<ElementId> linkElemIds = prjGridLevel.GetMonitoredLinkElementIds();
                    if (!linkElemIds.Any())
                    {
                        Print($"На элемент {prjGridLevel.Name} id: {prjGridLevel.Id} не назначен мониторинг. Оси/уровни - должны быть с мониторингом",
                            MessageType.Error);
                        return Result.Cancelled;
                    }
                    else if (linkElemIds.Count > 1)
                    {
                        Print($"На элемент {prjGridLevel.Name} id: {prjGridLevel.Id} мониторинг назначен более 1 раза. Это не допустимо",
                            MessageType.Error);
                        return Result.Cancelled;
                    }

                    string prjGridElemName = prjGridLevel.Name;
                    LinkGridLevelData equalLinkGridLevelData = linkGridLevelDataElemColl
                        .FirstOrDefault(lge => lge.Element.Name.Equals(prjGridElemName));
                    if (equalLinkGridLevelData.Equals(default))
                    {
                        // Кастомизация под проект
                        if (isSET && prjGridLevel is Level prjLvl)
                        {
                            equalLinkGridLevelData = linkGridLevelDataElemColl
                                .FirstOrDefault(lge => lge.Element is Level lgeLvl && (Math.Round(Math.Abs(lgeLvl.Elevation) - (Math.Abs(prjLvl.Elevation)), 1) == 0));
                        }
                        if (equalLinkGridLevelData.Equals(default))
                        {
                            Print(
                                $"Элемент {prjGridLevel.Name} id: {prjGridLevel.Id} - не существует в разб. файле. Проверь имена элементов вручную.",
                                MessageType.Error);
                            return Result.Cancelled;
                        }
                    }
                    else if (equalLinkGridLevelData.RevLinkInstance.Id != linkElemIds.FirstOrDefault())
                    {
                        Print($"Ошибка в назначении мониторинга для элемента {prjGridLevel.Name} id: {prjGridLevel.Id} - не из того файла скопирован элемент",
                            MessageType.Error);
                        return Result.Cancelled;
                    }
                    #endregion

                    #region Установка парамтера "На этаж выше"
                    if (prjGridLevel is Level prjLevel)
                    {
                        ElementId levelUpLevelId = new ElementId(-1);
                        Level equalLinkLevel = equalLinkGridLevelData.Element as Level;

                        ElementId equalLinkLevelUpParamData = equalLinkLevel.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId();
                        if (equalLinkLevelUpParamData != null
                            && equalLinkGridLevelData.RevLinkInstance.GetLinkDocument().GetElement(equalLinkLevelUpParamData) is Level equalLinkUpLevel)
                        {
                            Element levelUpLevelEqualLink = prjGridsElemColl.FirstOrDefault(pge => pge is Level && pge.Name.Equals(equalLinkUpLevel.Name));
                            if (levelUpLevelEqualLink == null)
                            {
                                // Кастомизация под проект
                                if (isSET)
                                {
                                    levelUpLevelEqualLink = prjGridsElemColl
                                        .FirstOrDefault(pge => pge is Level pgeLvl && (Math.Round(Math.Abs(pgeLvl.Elevation) - (Math.Abs(equalLinkUpLevel.Elevation)), 1) == 0));
                                }
                            }

                            if (levelUpLevelEqualLink == null)
                            {
                                Print(
                                    $"При поиске уровня сверху, уровень из связи {equalLinkUpLevel.Name} id: {equalLinkUpLevel.Id} - не нашел аналогичного в твоей модели. " +
                                    $"Это не останавливает анализ, и если у тебя выше этого уровня нет геометрии - можно проигнорировать данное предупреждение",
                                    MessageType.Warning);
                            }
                            else
                            {
                                levelUpLevelId = levelUpLevelEqualLink.Id;
                                Parameter levelUpParam = prjLevel.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL);
                                levelUpParam.Set(levelUpLevelId);
                            }
                        }
                    }
                    #endregion

                    #region Установка текстовых пар-в
                    Parameter sectionParam = prjGridLevel.LookupParameter(_sectParamName);
                    Parameter levelParam = prjGridLevel.LookupParameter(_levelParamName);
                    // Параметр секции ОБЯЗАТЕЛЕН на 100%. Уровень - только для сложных проектов
                    if (sectionParam == null)
                    {
                        Print($"Элемент {prjGridLevel.Name} id: {prjGridLevel.Id} - не имеет нужного параметра ({_sectParamName}) для сепарации на секции. Добавь его параметром проекта из ФОП_КПЛН",
                            MessageType.Error);
                        return Result.Cancelled;
                    }
                    
                    Parameter sectionParamLink = equalLinkGridLevelData.Element.LookupParameter(_sectParamName);
                    Parameter levelParamLink = equalLinkGridLevelData.Element.LookupParameter(_levelParamName);
                    // Параметр секции ОБЯЗАТЕЛЕН на 100%. Уровень - только для сложных проектов
                    if (sectionParamLink == null)
                    {
                        Print($"Элемент {equalLinkGridLevelData.Element.Name} из связи: {equalLinkGridLevelData.RevLinkInstance.Name} - не имеет нужного параметра ({_sectParamName}) для сепарации на секции. Добавь его параметром проекта из ФОП_КПЛН",
                            MessageType.Error);
                        return Result.Cancelled;
                    }
                    // Если у связи есть параметр уровня - то и у проекта он должен быть
                    if (levelParamLink != null && levelParam == null)
                    {
                        Print($"Элемент {prjGridLevel.Name} id: {prjGridLevel.Id} - не имеет нужного параметра ({_levelParamName}) для сепарации на этажи. Добавь его параметром проекта из ФОП_КПЛН",
                            MessageType.Error);
                        return Result.Cancelled;
                    }

                    // Обрабатываю секции
                    string sectionParamLinkData = sectionParamLink.AsString();
                    sectionParam.Set(sectionParamLinkData);
                    
                    if (!string.IsNullOrEmpty(sectionParamLinkData))
                        resultGridSections.Add(sectionParamLinkData);

                    // Обрабатываю этажи (если они есть)
                    if (levelParamLink == null) 
                        continue;
                    
                    string levelParamLinkData = levelParamLink.AsString();
                    levelParam.Set(levelParamLinkData);

                    if (!string.IsNullOrEmpty(levelParamLinkData))
                        resultLevels.Add(levelParamLinkData);
                    #endregion
                }

                userResultSectionsMsg = string.Join(", ", resultGridSections);
                userResultLevelsMsg = string.Join(", ", resultLevels);
                trans.Commit();
            }

            string mainContent;
            if (userResultLevelsMsg.Any())
                mainContent = $"Перенос данных из разбивочного файла выполнен успешно!" +
                    $"\nВ проекте выявлены форматы секций: {userResultSectionsMsg}" +
                    $"\nВ проекте выявлены форматы этажей: {userResultLevelsMsg}";
            else
                mainContent = $"Перенос данных из разбивочного файла выполнен успешно!" +
                    $"\nВ проекте выявлены форматы секций: {userResultSectionsMsg}";

            TaskDialog taskDialog = new TaskDialog("Результат работы")
            {
                MainContent = mainContent,
            };
            taskDialog.Show();

            return Result.Succeeded;
        }
    }

    // Контейнер для сеток связи
    internal struct LinkGridLevelData : IEquatable<LinkGridLevelData>
    {
        public Element Element { get; set; }
        public RevitLinkInstance RevLinkInstance { get; set; }

        public bool Equals(LinkGridLevelData other)
        {
            return Equals(Element, other.Element) && Equals(RevLinkInstance, other.RevLinkInstance);
        }

        public override bool Equals(object obj)
        {
            return obj is LinkGridLevelData other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Element != null ? Element.GetHashCode() : 0) * 397) ^ (RevLinkInstance != null ? RevLinkInstance.GetHashCode() : 0);
            }
        }
    }
}
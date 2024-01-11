using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal sealed class LevelAndGridsParamCopier : IExternalCommand
    {
        private readonly string _sectParamName = "КП_О_Секция";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            #region Получаю сетки из связей
            List<LinkGridData> linkGridDataElemColl = new List<LinkGridData>();
            RevitLinkInstance[] rvtLinkInstsColl = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                // Слабое место - имена файлов могут отличаться из-за требований Заказчика
                .Where(lm => lm.Name.ToUpper().Contains("_РАЗБ"))
                .Cast<RevitLinkInstance>()
                .ToArray();
            foreach (RevitLinkInstance rvtLinkInst in rvtLinkInstsColl)
            {
                Document linkDoc = rvtLinkInst.GetLinkDocument();
                linkGridDataElemColl.AddRange(
                    new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Grids)
                    .WhereElementIsNotElementType()
                    .Select(e => new LinkGridData() { Element = e, RevLinkInstance = rvtLinkInst }));
                //linkGridDataElemColl.AddRange(
                //    new FilteredElementCollector(linkDoc)
                //    .OfCategory(BuiltInCategory.OST_Levels)
                //    .WhereElementIsNotElementType()
                //    .Select(e => new LinkGridData() { Element = e, RevLinkInstance = rvtLinkInst }));
            }

            if (!linkGridDataElemColl.Any())
            {
                Print("Не удалось получить оси и уровни из разбивочного файла. Нужно загрузить файл в проект. Если файлы подгружены, то скинь проблему разработчику", 
                    KPLN_Loader.Preferences.MessageType.Error);
                return Result.Cancelled;
            }
            #endregion

            #region Получаю сетки из проекта
            IEnumerable<Element> prjGridsElemEnum = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();
            Element[] prjGridsElemColl = prjGridsElemEnum
                //.Concat(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                //.WhereElementIsNotElementType())
                .ToArray();
            #endregion

            string userResultMsg;
            HashSet<string> resultGridSections = new HashSet<string>();
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("КП: Копировать пар-ры сеток");

                foreach (Element prjGrid in prjGridsElemColl)
                {
                    #region Подготовка и проверка элемента
                    IList<ElementId> linkElemIds = prjGrid.GetMonitoredLinkElementIds();
                    if (!linkElemIds.Any())
                    {
                        Print($"На элемент {prjGrid.Name} id: {prjGrid.Id} не назначен мониторинг. Оси/уровни - должны быть с мониторингом",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }
                    else if (linkElemIds.Count > 1)
                    {
                        Print($"На элемент {prjGrid.Name} id: {prjGrid.Id} мониторинг назначен более 1 раза. Это не допустимо",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }

                    string prjGridElemName = prjGrid.Name;
                    LinkGridData equalLinkGridData = linkGridDataElemColl
                        .Where(lge => lge.Element.Name.Equals(prjGridElemName))
                        .FirstOrDefault();
                    if (equalLinkGridData.Equals(default(LinkGridData)))
                    {
                        Print($"Элемент {prjGrid.Name} id: {prjGrid.Id} - не существует в разб. файле. Проверь имена элементов вручную.",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }
                    else if (equalLinkGridData.RevLinkInstance.Id != linkElemIds.FirstOrDefault())
                    {
                        Print($"Ошибка в назначении мониторинга для элемента {prjGrid.Name} id: {prjGrid.Id} - не из того файла скопирован элемент",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }
                    #endregion

                    Parameter sectionParam = prjGrid.LookupParameter(_sectParamName);
                    if (sectionParam == null)
                    {
                        Print($"Элемент {prjGrid.Name} id: {prjGrid.Id} - не имеет нужного параметра ({_sectParamName}) для сепарации на секции. Добавь его параметром проекта из ФОП_КПЛН",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }
                    Parameter sectionParamLink = equalLinkGridData.Element.LookupParameter(_sectParamName);
                    if (sectionParamLink == null)
                    {
                        Print($"Элемент {equalLinkGridData.Element.Name} из связи: {equalLinkGridData.RevLinkInstance.Name} - не имеет нужного параметра ({_sectParamName}) для сепарации на секции. Добавь его параметром проекта из ФОП_КПЛН",
                            KPLN_Loader.Preferences.MessageType.Error);
                        return Result.Cancelled;
                    }

                    string sectionParamLinkData = sectionParamLink.AsString();
                    sectionParam.Set(sectionParamLinkData);

                    if (!string.IsNullOrEmpty(sectionParamLinkData))
                        resultGridSections.Add(sectionParamLinkData);
                }

                userResultMsg = string.Join(", ", resultGridSections);
                trans.Commit();
            }

            TaskDialog taskDialog = new TaskDialog("Результат работы")
            {
                MainContent = $"Перенос данных из разбивочного файла выполнен успешно!\nВ проекте выявлены форматы секций: {userResultMsg}",
            };
            taskDialog.Show();

            return Result.Succeeded;
        }
    }

    // Контейнер для сеток связи
    internal struct LinkGridData
    {
        public Element Element { get; set; }
        public RevitLinkInstance RevLinkInstance { get; set; }
    }
}
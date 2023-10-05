using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckListAnnotations : AbstrCheckCommand, IExternalCommand
    {
        /// <summary>
        /// Список категорий для анализа
        /// </summary>
        private readonly List<BuiltInCategory> _bicErrorSearch = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_Lines,
            BuiltInCategory.OST_TextNotes,
            BuiltInCategory.OST_RasterImages,
            BuiltInCategory.OST_GenericAnnotation,
            BuiltInCategory.OST_DetailComponents
        };

        /// <summary>
        /// Список исключений в именах семейств для генерации исключений в выбранных категориях
        /// </summary>
        private readonly List<string> _exceptionFamilyNameList = new List<string>
        {
            "011_",
            "012_",
            "020_Эквив",
            "022_",
            "023_",
            "024_",
            "070_",
            "099_"
        };

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            _name = "Проверка листов на аннотации";
            _application = uiapp;

            _allStorageName = "KPLN_CheckAnnotation";
            
            _lastRunGuid = new Guid("caf1c9b7-14cc-4ba1-8336-aa4b347d2898");
            _userTextGuid = new Guid("caf1c9b7-14cc-4ba1-8336-aa4b347d2899");

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;

            // Обрабатываю пользовательскую выборку листов
            List<ViewSheet> sheetsList = new List<ViewSheet>();
            List<ElementId> selIds = uidoc.Selection.GetElementIds().ToList();
            if (selIds.Count > 0)
            {
                foreach (ElementId selId in selIds)
                {
                    Element elem = doc.GetElement(selId);
                    int catId = elem.Category.Id.IntegerValue;
                    if (catId.Equals((int)BuiltInCategory.OST_Sheets))
                    {
                        ViewSheet curViewSheet = elem as ViewSheet;
                        sheetsList.Add(curViewSheet);
                    }
                }
                if (sheetsList.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "В выборке нет ни одного листа :(", TaskDialogCommonButtons.Ok);
                    return Result.Cancelled;
                }
            }

            #region Проверяю и обрабатываю элементы
            try
            {
                // Анализирую выбранные листы на количество аннотаций (с общим выводом в спец. окно)
                if (sheetsList.Count > 0)
                {
                    List<WPFEntity> wpfColl = new List<WPFEntity>();
                    foreach (ViewSheet viewSheet in sheetsList)
                    {
                        IEnumerable<WPFEntity> annColl = CheckCommandRunner(doc, FindAllAnnotationsOnList(doc, viewSheet));
                        if (annColl != null)
                            wpfColl.AddRange(annColl);
                    }
                    OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
                    if (form != null)
                    {
                        form.Show();
                        return Result.Succeeded;
                    }
                }

                // Анализирую все видовые экраны активного листа
                else if (activeView.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Sheets))
                {
                    ViewSheet viewSheet = activeView as ViewSheet;
                    IEnumerable<Element> annColl = FindAllAnnotationsOnList(doc, viewSheet);
                    QuickShowResult(uidoc, annColl);

                    // Провожу фиксацию запуска отдельно от вшитого в OutputMainForm
                    ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESBuilderRun, DateTime.Now));
                }

                // Анализирую вид
                else
                {
                    IEnumerable<Element> annColl = FindAllAnnotations(doc, activeView.Id);
                    QuickShowResult(uidoc, annColl);

                    // Провожу фиксацию запуска отдельно от вшитого в OutputMainForm
                    ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESBuilderRun, DateTime.Now));
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.InnerException.Message} \nStackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Header);
                else
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.Message} \nStackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Header);

                return Result.Cancelled;
            }
            #endregion

            return Result.Succeeded;
        }

        private protected override List<CheckCommandError> CheckElements(Document doc, IEnumerable<Element> elemColl) => null;

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, IEnumerable<Element> elemColl)
        {
            if (elemColl.Any())
            {
                List<WPFEntity> result = new List<WPFEntity>();

                string wpfInfoRow = string.Empty;
                foreach (Element elem in elemColl)
                {
                    if (doc.GetElement(elem.OwnerViewId) is ViewSheet viewSheet)
                    {
                        wpfInfoRow = $"Лист: {viewSheet.SheetNumber} - {viewSheet.Name}";
                        break;
                    }
                    else if (doc.GetElement(elem.OwnerViewId) is View view)
                    {
                        Parameter listNumberParam = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER);
                        Parameter listNameParam = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NAME);
                        if (listNumberParam == null || listNumberParam == null) wpfInfoRow = $"Вид: {view.Name}";
                        else wpfInfoRow = $"Лист: {listNumberParam.AsString()} - {listNameParam.AsString()}";
                        break;
                    }
                }

                result.Add(new WPFEntity(
                    elemColl,
                    Status.Error,
                    "Недопустимые аннотации",
                    "Данные элементы запрещено использовать на моделируемых видах",
                    false,
                    false,
                    wpfInfoRow));

                return result;
            }

            return null;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByStatus();
        }

        /// <summary>
        /// Метод для создания фильтра, для игнорирования элементов по имени семейства (НАЧИНАЕТСЯ С)
        /// </summary>
        private FilteredElementCollector FilteredByNotBeginsStringColl(FilteredElementCollector currentColl)
        {
            foreach (string currentName in _exceptionFamilyNameList)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                currentColl.WherePasses(eFilter);
            }
            return currentColl;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на единице выбранного элемента и записи в коллекцию
        /// </summary>
        private IEnumerable<Element> FindAllAnnotations(Document doc, ElementId viewId)
        {
            List<Element> result = new List<Element>();

            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(doc, viewId).OfCategory(bic).WhereElementIsNotElementType();
                // Очищаю коллецию от DetailLine - Последовательность компонентов узлов. Используется для разных визуальных маскировок, или докрутки видимости (УГО)
                result.AddRange(FilteredByNotBeginsStringColl(bicColl).ToElements().Where(e => e.GetType().Name != nameof(DetailLine)));
            }

            return result;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на листах и записи в коллекцию или словарь (в зависимости от количества выбранных листов)
        /// </summary>
        private IEnumerable<Element> FindAllAnnotationsOnList(Document doc, ViewSheet viewSheet)
        {
            List<Element> result = new List<Element>();

            // Анализирую аннотации на листе
            result.AddRange(FindAllAnnotations(doc, viewSheet.Id));

            // Анализирую размещенные виды
            ICollection<ElementId> allViewPorts = viewSheet.GetAllViewports();
            foreach (ElementId vpId in allViewPorts)
            {
                Viewport vp = (Viewport)doc.GetElement(vpId);
                ElementId viewId = vp.ViewId;
                Element currentElement = doc.GetElement(viewId);
                // Анализирую все виды, кроме чертежных видов и легенд
                if (!currentElement.GetType().Equals(typeof(ViewDrafting)) & !currentElement.GetType().Equals(typeof(View)))
                {
                    result.AddRange(FindAllAnnotations(doc, viewId));
                }
            }

            return result;
        }

        /// <summary>
        /// Метод для вывода микроотчета пользователю, а также для выделения элементов в модели
        /// </summary>
        private void QuickShowResult(UIDocument uidoc, IEnumerable<Element> elemColl)
        {
            if (elemColl.Count() == 0) TaskDialog.Show("Результат", "Аннотации не обнаружены :)", TaskDialogCommonButtons.Ok);
            else TaskDialog.Show("Результат", $"Аннотации выделены. Количество - {elemColl.Count()}", TaskDialogCommonButtons.Ok);

            // Выделяю элементы в модели
            uidoc.Selection.SetElementIds(elemColl.Select(e => e.Id).ToList());
        }
    }
}

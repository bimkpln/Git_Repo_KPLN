using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckListAnnotations : AbstrCheck
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
            "011_",
            "012_",
            "020_Эквив",
            "022_",
            "020_Эквив",
            "023_",
            "024_",
            "030_Проем",
            "070_",
            "071_Фон арм",
            "076_Маркировка позиций арм",
            "099_"
        };

        public CheckListAnnotations() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка на аннотации";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckAnnotation",
                    new Guid("caf1c9b7-14cc-4ba1-8336-aa4b357d2898"));
        }

        /// <summary>
        /// В данном случае - получаю анализируемые виды/листы
        /// </summary>
        /// <returns></returns>
        public override Element[] GetElemsToCheck()
        {
            List<Element> result = new List<Element>();

            // Обрабатываю пользовательскую выборку листов
            List<ElementId> selIds = CheckUIApp.ActiveUIDocument.Selection.GetElementIds().ToList();
            if (selIds.Count > 0)
            {
                foreach (ElementId selId in selIds)
                {
                    Element elem = CheckDocument.GetElement(selId);

#if Debug2020 || Revit2020
                    int catId = elem.Category.Id.IntegerValue;
                    if (catId.Equals((int)BuiltInCategory.OST_Sheets))
#else
                    if (elem.Category.BuiltInCategory == BuiltInCategory.OST_Sheets)
#endif
                    {
                        ViewSheet curViewSheet = elem as ViewSheet;
                        result.Add(curViewSheet);
                    }
                }

                if (result.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "В выборке нет ни одного листа :(", TaskDialogCommonButtons.Ok);
                    return new Element[0];
                }

            }

            // Анализирую все видовые экраны активного листа
#if Debug2020 || Revit2020
            else if (CheckUIApp.ActiveUIDocument.ActiveView.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Sheets))
#else
            else if (CheckUIApp.ActiveUIDocument.ActiveView.Category.BuiltInCategory == BuiltInCategory.OST_Sheets)
#endif
            {
                if (CheckUIApp.ActiveUIDocument.ActiveView is ViewSheet viewSheet)
                {
                    result.Add(viewSheet);
                    OnlySelectOnModel = true;
                }
                else
                {
                    TaskDialog.Show("Ошибка", "Активный вид не является ViewSheet.\n\nОтправь разработчику!", TaskDialogCommonButtons.Ok);
                    return new Element[0];
                }
            }

            // Анализирую вид
            else
            {
                result.Add(CheckUIApp.ActiveUIDocument.ActiveView);
                OnlySelectOnModel = true;
            }


            return result.ToArray();
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            foreach (Element elem in elemColl)
            {
                // Обрабатываю листы
                if (elem is ViewSheet viewSheet)
                {
                    CheckerEntity checkListAnnot = PrepareEntitiesOnList(viewSheet);
                    if (checkListAnnot != null)
                        _checkerEntitiesCollHeap.Add(checkListAnnot);
                }
                // Отрабатываю виды (т.е. всё остальное)
                else
                {
                    Element cElView = CheckDocument.GetElement(CheckUIApp.ActiveUIDocument.ActiveView.Id);
                    if (cElView is View currentView)
                    {
                        CheckerEntity checkViewAnnot = PrepareEntitiesOnView(currentView);
                        if (checkViewAnnot != null)
                            _checkerEntitiesCollHeap.Add(checkViewAnnot);
                    }
                }
            }

            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на листах
        /// </summary>
        private CheckerEntity PrepareEntitiesOnList(ViewSheet viewSheet)
        {
            // Анализирую аннотации на листе
            List<Element> annotations = FindAllAnnotations(viewSheet.Id).ToList();

            // Анализирую размещенные виды
            ICollection<ElementId> allViewPorts = viewSheet.GetAllViewports();
            foreach (ElementId vpId in allViewPorts)
            {
                Viewport vp = (Viewport)CheckDocument.GetElement(vpId);
                ElementId viewId = vp.ViewId;
                Element currentElement = CheckDocument.GetElement(viewId);

                // Анализирую все виды, кроме чертежных видов и легенд
                if (!currentElement.GetType().Equals(typeof(ViewDrafting)) & !currentElement.GetType().Equals(typeof(View)))
                    annotations.AddRange(FindAllAnnotations(viewId));
            }

            // Формирую сущности
            if (annotations.Any())
                return new CheckerEntity(
                    annotations,
                    "Недопустимые аннотации",
                    "Данные элементы запрещено использовать на моделируемых видах",
                    $"Лист: {viewSheet.SheetNumber} - {viewSheet.Name}");

            return null;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на виде
        /// </summary>
        private CheckerEntity PrepareEntitiesOnView(View view)
        {
            // Получаю аннотации на виде
            List<Element> annotations = new List<Element>();
            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(CheckDocument, view.Id).OfCategory(bic).WhereElementIsNotElementType();
                annotations.AddRange(FilteredByNotBeginsStringColl(bicColl).ToElements());
            }

            // Формирую сущности
            if (annotations.Any())
                return new CheckerEntity(
                    annotations,
                    "Недопустимые аннотации",
                    "Данные элементы запрещено использовать на моделируемых видах",
                    $"Вид: {view.Name}");

            return null;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на единице выбранного элемента и записи в коллекцию
        /// </summary>
        private IEnumerable<Element> FindAllAnnotations(ElementId viewId)
        {
            List<Element> result = new List<Element>();

            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(CheckDocument, viewId).OfCategory(bic).WhereElementIsNotElementType();
                result.AddRange(FilteredByNotBeginsStringColl(bicColl).ToElements());
                // Очищаю коллецию от DetailLine - Последовательность компонентов узлов. Используется для разных визуальных маскировок, или докрутки видимости (УГО)
                //result.AddRange(FilteredByNotBeginsStringColl(bicColl).ToElements().Where(e => e.GetType().Name != nameof(DetailLine)));
            }

            return result;
        }

        /// <summary>
        /// Метод для создания фильтра, для игнорирования элементов по имени семейства (НАЧИНАЕТСЯ С)
        /// </summary>
        private FilteredElementCollector FilteredByNotBeginsStringColl(FilteredElementCollector currentColl)
        {
            foreach (string currentName in _exceptionFamilyNameList)
            {
#if Debug2020 || Revit2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName);
#endif

                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                currentColl.WherePasses(eFilter);
            }

            return currentColl;
        }
    }
}

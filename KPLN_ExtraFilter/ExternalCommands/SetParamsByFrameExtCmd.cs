using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    /// <summary>
    /// Класс фильтрации Selection 
    /// </summary>
    internal class SelectorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Фильтрация по классу (проемы)
            if (elem is Opening _)
                return false;

            // Фильтрация по категории (модельные эл-ты, кроме видов, сборок)
            if (elem.Category is Category elCat)
            {
                int elCatId = elCat.Id.IntegerValue;
                if (((elem.Category.CategoryType == CategoryType.Model)
                        || (elem.Category.CategoryType == CategoryType.Internal))
                    && (elCatId != (int)BuiltInCategory.OST_Viewers)
                    && (elCatId != (int)BuiltInCategory.OST_IOSModelGroups)
                    && (elCatId != (int)BuiltInCategory.OST_Assemblies))
                    return true;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SetParamsByFrameExtCmd : IExternalCommand
    {
        internal const string PluginName = "Выбрать/заполнить рамкой";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                IList<Reference> selectionRefers = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new SelectorFilter(),
                    "Выберите нужные элементы (рамкой, по одному) и нажмите \"Готово\"");

                #region Подготовка параметров для запуска
                // Выделенные эл-ты
                Element[] selectedElemsToFind = selectionRefers
                    .Select(r => doc.GetElement(r.ElementId))
                    .ToArray();

                if (!selectedElemsToFind.Any())
                    return Result.Cancelled;

                // Расширенная выборка к выделенным
                Element[] expandedElemsToFind = ExtraSelection(doc, selectedElemsToFind).ToArray();

                // Чистка коллекции от экз. Одинаковых семейств
                // (многопоточность не справиться из-за ревит, поэтому нужно предв. Очистка)
                Element[] clearedElemsToFind = expandedElemsToFind
                    .GroupBy(x => x.GetTypeId())
                    .Select(gr => gr.FirstOrDefault())
                    .ToArray();

                Parameter[] elemsParams = DocWorker.GetUnionParamsFromElems(doc, clearedElemsToFind).ToArray();

                List<ParamEntity> allParamsEntities = new List<ParamEntity>(elemsParams.Count());
                foreach (Parameter param in elemsParams)
                {
                    string toolTip = string.Empty;
                    if (param.IsShared)
                        toolTip = $"Id: {param.Id}, GUID: {param.GUID}";
                    else if (param.Id.IntegerValue < 0)
                        toolTip = $"Id: {param.Id}, это СИСТЕМНЫЙ параметр проекта";
                    else
                        toolTip = $"Id: {param.Id}, это ПОЛЬЗОВАТЕЛЬСКИЙ параметр проекта";

                    allParamsEntities.Add(new ParamEntity(param, toolTip));
                }
                #endregion

                // Чтение конфигурации последнего запуска
                object lastRunConfigObj = ConfigService.ReadConfigFile<List<MainItem>>(ModuleData.RevitVersion, doc, ConfigType.Memory);

                // Подготовка ViewModel для старта окна
                SetParamsByFrameForm form = new SetParamsByFrameForm(expandedElemsToFind, allParamsEntities, lastRunConfigObj);
                form.ShowDialog();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Расширенное выделение элементов модели
        /// </summary>
        /// <param name="doc">Ревит-док</param>
        /// <param name="selectedElems">Коллекция выделенных в ревит эл-в</param>
        /// <returns></returns>
        private static IEnumerable<Element> ExtraSelection(Document doc, Element[] selectedElems)
        {
            List<Element> result = new List<Element>(selectedElems);

            ElementClassFilter famIsntFilter = new ElementClassFilter(typeof(FamilyInstance));

            // Изоляция воздуховодов и труб выделяется рамкой селектора
            List<ElementFilter> filters = new List<ElementFilter>()
            {
                famIsntFilter,
            };
            LogicalOrFilter resultFilter = new LogicalOrFilter(filters);

            foreach (Element elem in selectedElems)
            {
                IList<ElementId> depElems = elem.GetDependentElements(resultFilter);
                foreach (ElementId id in depElems)
                {
                    Element currentElem = doc.GetElement(id);
                    if (currentElem.Id.IntegerValue == elem.Id.IntegerValue)
                        continue;

                    // Игнорирую балясины (отдельно в спеки не идут)
                    if (currentElem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StairsRailingBaluster)
                        continue;

                    // Предварительно фильтрую общие вложенные семейства
                    if (currentElem is FamilyInstance famInst && famInst.SuperComponent != null)
                        result.Add(famInst);
                    // Добавляю ВСЕ остальное
                    else
                        result.Add(currentElem);
                }

            }

            return result;
        }
    }
}

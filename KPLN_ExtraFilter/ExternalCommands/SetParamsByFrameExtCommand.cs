using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.Forms;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    /// <summary>
    /// Класс для сравнения парамтеров
    /// </summary>
    internal class ParameterComparer : IEqualityComparer<Parameter>
    {
        public bool Equals(Parameter x, Parameter y) => x.Definition.Name == y.Definition.Name;

        public int GetHashCode(Parameter obj) => obj.Definition.Name.GetHashCode();
    }

    /// <summary>
    /// Класс фильтрации Selection 
    /// </summary>
    internal class SelectorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category != null
                && (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_IOSModelGroups)
                && (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Assemblies))
                return true;

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SetParamsByFrameExtCommand : IExternalCommand
    {
        /// <summary>
        /// Кэширование конфига предыдущего запуска
        /// </summary>
        public static string MemoryConfigData;
        
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
                
                // Расширенная выборка к выделенным
                Element[] expandedElemsToFind = ExtraSelection(doc, selectedElemsToFind).ToArray();

                // Чистка коллекции от экз. Одинаковых семейств
                // (многопоточность не справиться из-за ревит, поэтому нужно предв. Очистка)
                Element[] clearedElemsToFind = expandedElemsToFind
                    .GroupBy(x => x.GetTypeId())
                    .Select(gr => gr.FirstOrDefault())
                    .ToArray();

                Parameter[] elemsParams = GetParamsFromElems(doc, clearedElemsToFind).ToArray();

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

                // Подготовка ViewModel для старта окна
                SetParamsByFrameForm form = null;
                if (string.IsNullOrEmpty(MemoryConfigData))
                    form = new SetParamsByFrameForm(expandedElemsToFind, allParamsEntities);
                else
                {
                    var jsonEnts = new ObservableCollection<MainItem>(
                        JsonConvert.DeserializeObject<List<MainItem>>(MemoryConfigData));

                    // Если такой пар-р есть в общей коллекции, значит добавляю его в форму. Иначе - нет
                    IEnumerable<MainItem> userSelectedViewModels = jsonEnts
                        .Where(entity => allParamsEntities.Count(ent => ent.CurrentParamIntId == entity.UserSelectedParamEntity.CurrentParamIntId) == 1)
                        .Select(entity =>
                        {
                            entity.UserSelectedParamEntity.CurrentParamName = allParamsEntities
                            .First(ent => ent.CurrentParamIntId == entity.UserSelectedParamEntity.CurrentParamIntId)
                            .CurrentParamName;
                            
                            return entity;
                        });

                    form = new SetParamsByFrameForm(expandedElemsToFind, allParamsEntities, userSelectedViewModels);
                }
                
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
        /// Получить парамеры из указанных элементов
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="elemsToFind">Коллекция элементов для анализа</param>
        /// <returns></returns>
        private static IEnumerable<Parameter> GetParamsFromElems(Document doc, Element[] elemsToFind)
        {
            // Основной блок
            HashSet<Parameter> commonInstParameters = new HashSet<Parameter>(new ParameterComparer());
            HashSet<Parameter> commonTypeParameters = new HashSet<Parameter>(new ParameterComparer());
            Element firstElement = elemsToFind.FirstOrDefault();

            AddParam(firstElement, commonInstParameters);

            Element typeElem = doc.GetElement(firstElement.GetTypeId());
            AddParam(typeElem, commonTypeParameters);

            foreach (Element currentElement in elemsToFind)
            {
                // Игнорирую уже добавленный эл-т
                if (firstElement.Id == currentElement.Id)
                    continue;

                HashSet<Parameter> currentInstParameters = new HashSet<Parameter>(new ParameterComparer());
                AddParam(currentElement, currentInstParameters);
                commonInstParameters.IntersectWith(currentInstParameters);

                HashSet<Parameter> currentTypeParameters = new HashSet<Parameter>(new ParameterComparer());
                Element currentTypeElem = doc.GetElement(currentElement.GetTypeId());
                AddParam(currentTypeElem, currentTypeParameters);
                commonTypeParameters.IntersectWith(currentTypeParameters);
            }

            if (commonInstParameters.Count == 0 && commonTypeParameters.Count == 0)
                throw new Exception("Ошибка в поиске парамеров для элементов Revit");

            return new HashSet<Parameter>(commonInstParameters.Union(commonTypeParameters), 
                new ParameterComparer());
        }

        /// <summary>
        /// Добавить пар-р в коллекцию с пред. подготовкой
        /// </summary>
        /// <param name="elem">Елемент для аналища</param>
        /// <param name="setToAdd">Коллекция для добавления</param>
        private static void AddParam(Element elem, HashSet<Parameter> setToAdd)
        {
            foreach (Parameter param in elem.Parameters)
            {
                // Отбрасываю системные пар-ры, которые нельзя редачить (Категория, Имя типа и т.п.) 
                if (param.Id.IntegerValue < 0 && param.IsReadOnly)
                    continue;
                
                setToAdd.Add(param);
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

            ElementClassFilter pipeInsulFilter = new ElementClassFilter(typeof(PipeInsulation));
            ElementClassFilter ductInsulFilter = new ElementClassFilter(typeof(DuctInsulation));
            ElementClassFilter famIsntFilter = new ElementClassFilter(typeof(FamilyInstance));

            List<ElementFilter> filters = new List<ElementFilter>()
            {
                pipeInsulFilter,
                ductInsulFilter,
                famIsntFilter,
            };
            LogicalOrFilter resultFilter = new LogicalOrFilter(filters);

            foreach (Element elem in selectedElems)
            {
                IList<ElementId> depElems = elem.GetDependentElements(resultFilter);
                foreach(ElementId id in depElems)
                {
                    Element currentElem = doc.GetElement(id);
                    if (currentElem.Id.IntegerValue == elem.Id.IntegerValue)
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

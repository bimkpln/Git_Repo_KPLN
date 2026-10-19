using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB.Events;

namespace KPLN_ViewsAndLists_Ribbon.Services.Schedules
{
    /// <summary>Поля штатного списка Revit, без неиспользуемых общих параметров проекта.</summary>
    internal sealed class ScheduleParameterCatalog
    {
        private readonly Document _document;
        private readonly Dictionary<ElementId, HashSet<ElementId>> _available = new Dictionary<ElementId, HashSet<ElementId>>();
        private readonly Dictionary<ElementId, HashSet<ElementId>> _bound = new Dictionary<ElementId, HashSet<ElementId>>();
        private static readonly ConditionalWeakTable<Document, FamilyCache> FamilyCaches = new ConditionalWeakTable<Document, FamilyCache>();
        private static readonly ConditionalWeakTable<Autodesk.Revit.ApplicationServices.Application, object> SubscribedApplications = new ConditionalWeakTable<Autodesk.Revit.ApplicationServices.Application, object>();

        private sealed class FamilyCache
        {
            // Только GUID: объекты Revit не переживают перезагрузку семейства или откат.
            public readonly Dictionary<string, Guid[]> Parameters = new Dictionary<string, Guid[]>();
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            Document document = args.GetDocument();
            FamilyCache cache;
            if (!FamilyCaches.TryGetValue(document, out cache)) return;

            // Удалённый объект уже недоступен для определения класса — сбрасываем консервативно.
            if (args.GetDeletedElementIds().Count > 0 ||
                args.GetAddedElementIds().Concat(args.GetModifiedElementIds()).Any(id =>
                {
                    Element element = document.GetElement(id);
                    return element is Family || element is FamilySymbol || element is ParameterElement;
                }))
                cache.Parameters.Clear();
        }

        public ScheduleParameterCatalog(Document document)
        {
            _document = document;
            SubscribedApplications.GetValue(document.Application, application =>
            {
                application.DocumentChanged += OnDocumentChanged;
                return new object();
            });
            var iterator = document.ParameterBindings.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                var parameter = iterator.Key as InternalDefinition;
                var binding = iterator.Current as ElementBinding;
                if (parameter == null || binding == null) continue;

                foreach (Category category in binding.Categories)
                {
                    HashSet<ElementId> ids;
                    if (!_bound.TryGetValue(category.Id, out ids))
                        _bound.Add(category.Id, ids = new HashSet<ElementId>());
                    ids.Add(parameter.Id);
                }
            }
        }

        public bool Contains(ScheduleDefinition definition, ElementId parameterId, ScheduleFieldType fieldType)
        {
            // Системные поля уже отобраны GetSchedulableFields для этой спецификации.
            // TableView не возвращает часть из них (например, семейство и тип).
#if Debug2024 || Revit2024
            if (parameterId.Value < -1) return true;
#else
            if (parameterId.IntegerValue < -1) return true;
#endif

            ElementId categoryId = SourceCategory(definition.CategoryId, fieldType);
            HashSet<ElementId> ids;
            if (!_available.TryGetValue(categoryId, out ids))
            {
                ids = new HashSet<ElementId>(TableView.GetAvailableParameters(_document, categoryId));
                HashSet<ElementId> bound;
                if (_bound.TryGetValue(categoryId, out bound)) ids.UnionWith(bound);
                AddFamilyParameters(categoryId, ids);
                _available.Add(categoryId, ids);
            }

            return ids.Contains(parameterId);
        }

        private void AddFamilyParameters(ElementId categoryId, HashSet<ElementId> ids)
        {
            FamilyCache cache = FamilyCaches.GetValue(_document, document => new FamilyCache());
            // Один представитель семейства: набор параметров не зависит от числа его экземпляров.
            // Коллектор ограничен категорией выбранной спецификации, а не всей моделью.
            var represented = new HashSet<ElementId>();
            using (var collector = new FilteredElementCollector(_document).OfClass(typeof(FamilyInstance)))
            {
                if (categoryId != ElementId.InvalidElementId) collector.OfCategoryId(categoryId);
                foreach (FamilyInstance instance in collector)
                {
                    FamilySymbol symbol = instance.Symbol;
                    if (symbol == null || !represented.Add(symbol.Family.Id)) continue;
                    AddSharedParameters(instance, ids);
                    AddSharedParameters(symbol, ids);
                }
            }

            // Загруженные, но ещё не размещённые семейства тоже дают поля в штатном окне Revit.
            using (var collector = new FilteredElementCollector(_document).OfClass(typeof(Family)))
            {
                foreach (Family family in collector)
                {
                    if (family.FamilyCategory == null ||
                        (categoryId != ElementId.InvalidElementId && family.FamilyCategory.Id != categoryId) ||
                        represented.Contains(family.Id) || family.IsInPlace) continue;

                    Guid[] cached;
                    if (cache.Parameters.TryGetValue(family.UniqueId, out cached))
                    {
                        AddSharedParameters(cached, ids);
                        continue;
                    }

                    Document familyDocument = null;
                    bool opened = false;
                    var sharedGuids = new List<Guid>();
                    try
                    {
                        familyDocument = _document.Application.Documents.Cast<Document>()
                            .FirstOrDefault(doc => doc.IsFamilyDocument && doc.OwnerFamily.Name == family.Name);
                        if (familyDocument == null)
                        {
                            if (!family.IsEditable) continue;
                            familyDocument = _document.EditFamily(family);
                            opened = true;
                        }

                        foreach (FamilyParameter parameter in familyDocument.FamilyManager.Parameters)
                        {
                            if (!parameter.IsShared) continue;
                            sharedGuids.Add(parameter.GUID);
                        }
                    }
                    finally
                    {
                        if (opened && familyDocument != null) familyDocument.Close(false);
                    }
                    cached = sharedGuids.ToArray();
                    cache.Parameters[family.UniqueId] = cached;
                    AddSharedParameters(cached, ids);
                }
            }
        }

        private void AddSharedParameters(IEnumerable<Guid> guids, HashSet<ElementId> ids)
        {
            foreach (Guid guid in guids)
            {
                var shared = SharedParameterElement.Lookup(_document, guid);
                if (shared != null) ids.Add(shared.Id);
            }
        }

        private static void AddSharedParameters(Element element, HashSet<ElementId> ids)
        {
            foreach (Parameter parameter in element.Parameters)
                if (parameter.IsShared) ids.Add(parameter.Id);
        }

        public IList<SchedulableField> GetFields(ScheduleDefinition definition)
        {
            return definition.GetSchedulableFields()
                .Where(field => Contains(definition, field.ParameterId, field.FieldType)).ToList();
        }

        private static ElementId SourceCategory(ElementId categoryId, ScheduleFieldType fieldType)
        {
            // В штатном окне это список «Выбирать доступные поля из».
            switch (fieldType)
            {
                case ScheduleFieldType.Room:
                case ScheduleFieldType.FromRoom:
                case ScheduleFieldType.ToRoom:
                    return new ElementId(BuiltInCategory.OST_Rooms);
                case ScheduleFieldType.Space:
                    return new ElementId(BuiltInCategory.OST_MEPSpaces);
                case ScheduleFieldType.Material:
                case ScheduleFieldType.StructuralMaterial:
                    return new ElementId(BuiltInCategory.OST_Materials);
                case ScheduleFieldType.ProjectInfo:
                    return new ElementId(BuiltInCategory.OST_ProjectInformation);
                case ScheduleFieldType.RevitLinkInstance:
                case ScheduleFieldType.RevitLinkType:
                    return new ElementId(BuiltInCategory.OST_RvtLinks);
                default:
                    return categoryId;
            }
        }
    }
}

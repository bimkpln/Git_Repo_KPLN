using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CommandsWheel.Services
{
    internal static class SelectionCustomCommandService
    {
        internal const string SelectAllInstancesVisibleInViewId = "KPLN_CUSTOM_SELECT_ALL_INSTANCES_VISIBLE_IN_VIEW";
        internal const string SelectAllInstancesInProjectId = "KPLN_CUSTOM_SELECT_ALL_INSTANCES_IN_PROJECT";
        internal const string SelectAllInstancesInProjectIncludingLegendsId = "KPLN_CUSTOM_SELECT_ALL_INSTANCES_IN_PROJECT_INCLUDING_LEGENDS";

        internal static IEnumerable<RevitCommandInfo> GetCommands()
        {
            yield return CreateCommand(
                SelectAllInstancesVisibleInViewId,
                "Выбрать все экземпляры: видимые на виде",
                "Выбирает все экземпляры того же типа, видимые на активном виде.");

            yield return CreateCommand(
                SelectAllInstancesInProjectId,
                "Выбрать все экземпляры: во всем проекте",
                "Выбирает все экземпляры того же типа во всем проекте, кроме легенд.");

            yield return CreateCommand(
                SelectAllInstancesInProjectIncludingLegendsId,
                "Выбрать все экземпляры: во всем проекте, включая легенды",
                "Выбирает все экземпляры того же типа во всем проекте, включая легенды.");
        }

        internal static void AddCommands(List<RevitCommandInfo> commands)
        {
            if (commands == null)
            {
                return;
            }

            HashSet<string> existingIds = new HashSet<string>(
                commands
                    .Where(command => command != null && !string.IsNullOrWhiteSpace(command.Id))
                    .Select(command => command.Id),
                StringComparer.OrdinalIgnoreCase);

            foreach (RevitCommandInfo command in GetCommands())
            {
                if (existingIds.Add(command.Id))
                {
                    commands.Add(command);
                }
            }
        }

        internal static bool TryExecute(UIApplication app, string commandId)
        {
            SelectAllInstancesScope scope;
            if (!TryGetScope(commandId, out scope))
            {
                return false;
            }

            ExecuteSelectAllInstances(app, scope);
            return true;
        }

        private static RevitCommandInfo CreateCommand(string id, string name, string tooltip)
        {
            return new RevitCommandInfo
            {
                Id = id,
                Name = name,
                TabName = "KPLN",
                PanelName = "Выбор",
                Tooltip = tooltip,
                CanPost = true
            };
        }

        private static bool TryGetScope(string commandId, out SelectAllInstancesScope scope)
        {
            if (string.Equals(commandId, SelectAllInstancesVisibleInViewId, StringComparison.OrdinalIgnoreCase))
            {
                scope = SelectAllInstancesScope.VisibleInView;
                return true;
            }

            if (string.Equals(commandId, SelectAllInstancesInProjectId, StringComparison.OrdinalIgnoreCase))
            {
                scope = SelectAllInstancesScope.InProject;
                return true;
            }

            if (string.Equals(commandId, SelectAllInstancesInProjectIncludingLegendsId, StringComparison.OrdinalIgnoreCase))
            {
                scope = SelectAllInstancesScope.InProjectIncludingLegends;
                return true;
            }

            scope = SelectAllInstancesScope.VisibleInView;
            return false;
        }

        private static void ExecuteSelectAllInstances(UIApplication app, SelectAllInstancesScope scope)
        {
            UIDocument uidoc = app == null ? null : app.ActiveUIDocument;
            Document document = uidoc == null ? null : uidoc.Document;
            if (uidoc == null || document == null)
            {
                TaskDialog.Show("Команды", "Нет активного документа Revit.");
                return;
            }

            Element sourceElement = GetSourceElement(uidoc, document);
            if (sourceElement == null)
            {
                return;
            }

            ElementId typeId = GetTypeId(sourceElement);
            if (typeId == null || ElementId.InvalidElementId.Equals(typeId))
            {
                TaskDialog.Show("Команды", "Для выбранного элемента не удалось определить тип.");
                return;
            }

            List<ElementId> elementIds = CollectInstances(document, typeId, scope);
            if (elementIds.Count == 0)
            {
                TaskDialog.Show("Команды", "Экземпляры того же типа не найдены.");
                return;
            }

            uidoc.Selection.SetElementIds(elementIds);
        }

        private static Element GetSourceElement(UIDocument uidoc, Document document)
        {
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            ElementId selectedId = selectedIds == null ? null : selectedIds.FirstOrDefault();
            if (selectedId != null && !ElementId.InvalidElementId.Equals(selectedId))
            {
                return document.GetElement(selectedId);
            }

            try
            {
                Reference pickedReference = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    "Выберите элемент для команды \"Выбрать все экземпляры\"");

                if (pickedReference == null || ElementId.InvalidElementId.Equals(pickedReference.ElementId))
                {
                    return null;
                }

                uidoc.Selection.SetElementIds(new[] { pickedReference.ElementId });
                return document.GetElement(pickedReference.ElementId);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        private static ElementId GetTypeId(Element element)
        {
            if (element == null)
            {
                return null;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId != null && !ElementId.InvalidElementId.Equals(typeId))
            {
                return typeId;
            }

            ElementType elementType = element as ElementType;
            return elementType == null ? ElementId.InvalidElementId : elementType.Id;
        }

        private static List<ElementId> CollectInstances(Document document, ElementId typeId, SelectAllInstancesScope scope)
        {
            FilteredElementCollector collector = scope == SelectAllInstancesScope.VisibleInView
                ? new FilteredElementCollector(document, document.ActiveView.Id)
                : new FilteredElementCollector(document);

            List<ElementId> result = new List<ElementId>();

            foreach (Element element in collector.WhereElementIsNotElementType())
            {
                if (element == null || !typeId.Equals(element.GetTypeId()))
                {
                    continue;
                }

                if (scope == SelectAllInstancesScope.InProject && IsOwnedByLegendView(document, element))
                {
                    continue;
                }

                result.Add(element.Id);
            }

            return result;
        }

        private static bool IsOwnedByLegendView(Document document, Element element)
        {
            if (document == null || element == null || ElementId.InvalidElementId.Equals(element.OwnerViewId))
            {
                return false;
            }

            View ownerView = document.GetElement(element.OwnerViewId) as View;
            return ownerView != null && ownerView.ViewType == ViewType.Legend;
        }

        private enum SelectAllInstancesScope
        {
            VisibleInView,
            InProject,
            InProjectIncludingLegends
        }
    }
}
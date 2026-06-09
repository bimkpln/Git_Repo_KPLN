using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExecutableCommand
{
    internal class ApartmentExecutionReportActionController : IApartmentExecutionReportActionController
    {
        private readonly ApartmentExecutionReportActionHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public ApartmentExecutionReportActionController()
        {
            _handler = new ApartmentExecutionReportActionHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void RequestShowElement(ApartmentExecutionReportItem item)
        {
            if (item == null)
                return;

            _handler.PrepareShow(item.GetNavigationCandidates());
            _externalEvent.Raise();
        }

        public void RequestDeleteRemnants(ApartmentExecutionReportItem item)
        {
            if (item == null)
                return;

            _handler.PrepareDelete(item.DeletableElementIds != null
                ? item.DeletableElementIds.ToList()
                : new List<long>());
            _externalEvent.Raise();
        }

        public void RequestRestore2DFamily(ApartmentExecutionReportItem item)
        {
            if (item == null || item.Restore2DInfo == null)
                return;

            _handler.PrepareRestore2D(item.Restore2DInfo);
            _externalEvent.Raise();
        }
    }

    internal class ApartmentExecutionReportActionHandler : IExternalEventHandler
    {
        private const string ApartmentInstanceMarker = "[KPLN_APT_INSTANCE]";

        private enum ReportAction
        {
            None,
            ShowElement,
            DeleteRemnants,
            Restore2DFamily
        }

        private ReportAction _action = ReportAction.None;
        private List<long> _elementIds = new List<long>();
        private Apartment2DRestoreInfo _restoreInfo;

        public string GetName()
        {
            return "KPLN. Действия отчёта квартир";
        }

        public void PrepareShow(List<long> elementIds)
        {
            _action = ReportAction.ShowElement;
            _elementIds = elementIds != null ? elementIds.Where(x => x > 0).Distinct().ToList() : new List<long>();
            _restoreInfo = null;
        }

        public void PrepareDelete(List<long> elementIds)
        {
            _action = ReportAction.DeleteRemnants;
            _elementIds = elementIds != null ? elementIds.Where(x => x > 0).Distinct().ToList() : new List<long>();
            _restoreInfo = null;
        }

        public void PrepareRestore2D(Apartment2DRestoreInfo restoreInfo)
        {
            _action = ReportAction.Restore2DFamily;
            _elementIds = new List<long>();
            _restoreInfo = restoreInfo;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                UIDocument uidoc = app.ActiveUIDocument;
                if (uidoc == null || uidoc.Document == null)
                    return;

                Document doc = uidoc.Document;

                switch (_action)
                {
                    case ReportAction.ShowElement:
                        ExecuteShowElement(uidoc, doc);
                        break;

                    case ReportAction.DeleteRemnants:
                        ExecuteDeleteRemnants(doc);
                        break;

                    case ReportAction.Restore2DFamily:
                        ExecuteRestore2DFamily(uidoc, doc);
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Отчёт квартир", ex.Message);
            }
            finally
            {
                _action = ReportAction.None;
                _elementIds = new List<long>();
                _restoreInfo = null;
            }
        }

        private void ExecuteShowElement(UIDocument uidoc, Document doc)
        {
            foreach (long idValue in _elementIds)
            {
                ElementId id = IDHelper.CreateElementId(idValue);
                Element element = doc.GetElement(id);
                if (element == null)
                    continue;

                uidoc.Selection.SetElementIds(new List<ElementId> { id });
                uidoc.ShowElements(id);
                return;
            }

            TaskDialog.Show("KPLN. Отчёт квартир", "Не найден ни один связанный элемент квартиры в модели.");
        }

        private void ExecuteDeleteRemnants(Document doc)
        {
            List<ElementId> ids = _elementIds
                .Select(IDHelper.CreateElementId)
                .Where(x => x != null && x != ElementId.InvalidElementId && doc.GetElement(x) != null)
                .ToList();

            if (ids.Count == 0)
            {
                TaskDialog.Show("KPLN. Отчёт квартир", "Элементы для удаления не найдены.");
                return;
            }

            int deletedCount = 0;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Удаление остатков квартиры"))
            {
                t.Start();

                foreach (ElementId id in ids)
                {
                    try
                    {
                        if (doc.GetElement(id) == null)
                            continue;

                        doc.Delete(id);
                        deletedCount++;
                    }
                    catch
                    {
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("KPLN. Отчёт квартир", "Удалено элементов: " + deletedCount);
        }

        private void ExecuteRestore2DFamily(UIDocument uidoc, Document doc)
        {
            if (_restoreInfo == null || !_restoreInfo.CanRestore)
            {
                TaskDialog.Show("KPLN. Отчёт квартир", "Нет данных для восстановления 2D-семейства.");
                return;
            }

            FamilySymbol symbol = doc.GetElement(IDHelper.CreateElementId(_restoreInfo.SymbolId)) as FamilySymbol;
            ViewPlan viewPlan = doc.GetElement(IDHelper.CreateElementId(_restoreInfo.ViewId)) as ViewPlan;

            if (symbol == null || viewPlan == null)
            {
                TaskDialog.Show("KPLN. Отчёт квартир", "Не найден тип 2D-семейства или вид для восстановления.");
                return;
            }

            FamilyInstance restored;
            XYZ point = new XYZ(_restoreInfo.X, _restoreInfo.Y, _restoreInfo.Z);

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Возврат 2D-семейства квартиры"))
            {
                t.Start();

                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }

                restored = CreateFamilyInstance(doc, viewPlan, symbol, point, _restoreInfo.LevelId);

                if (restored == null)
                    throw new Exception("Не удалось создать 2D-семейство.");

                if (Math.Abs(_restoreInfo.Rotation) > 1e-9)
                {
                    Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, restored.Id, axis, _restoreInfo.Rotation);
                }

                AppendComment(restored, ApartmentInstanceMarker);

                t.Commit();
            }

            uidoc.Selection.SetElementIds(new List<ElementId> { restored.Id });
            uidoc.ShowElements(restored.Id);
        }

        private static FamilyInstance CreateFamilyInstance(Document doc, ViewPlan viewPlan, FamilySymbol symbol, XYZ point, long levelIdValue)
        {
            FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;

            switch (placementType)
            {
                case FamilyPlacementType.ViewBased:
                    return doc.Create.NewFamilyInstance(point, symbol, viewPlan);

                case FamilyPlacementType.OneLevelBased:
                case FamilyPlacementType.OneLevelBasedHosted:
                case FamilyPlacementType.WorkPlaneBased:
                    Level level = null;
                    if (levelIdValue > 0)
                        level = doc.GetElement(IDHelper.CreateElementId(levelIdValue)) as Level;

                    if (level == null)
                        level = viewPlan.GenLevel;

                    if (level == null)
                        throw new Exception("У вида не определён уровень.");

                    return doc.Create.NewFamilyInstance(point, symbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                default:
                    throw new NotSupportedException("Тип размещения семейства не поддерживается: " + placementType);
            }
        }

        private static void AppendComment(Element element, string textToAppend)
        {
            if (element == null || string.IsNullOrWhiteSpace(textToAppend))
                return;

            Parameter p = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p == null || p.IsReadOnly)
                return;

            string oldValue = p.AsString();

            if (string.IsNullOrWhiteSpace(oldValue))
            {
                p.Set(textToAppend);
                return;
            }

            if (oldValue.Contains(textToAppend))
                return;

            p.Set(oldValue.TrimEnd() + " " + textToAppend);
        }

    }
}
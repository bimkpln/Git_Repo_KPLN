using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_ExtraFilter.Forms.Entities.SearchById;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_ExtraFilter.ExecutableCommand
{
    internal sealed class SelectByIdExсCmd : IExecutableCommand
    {
        private readonly SearchByIdEntity _searchEnt;

        public SelectByIdExсCmd(SearchByIdEntity searchEnt)
        {
            _searchEnt = searchEnt;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            try
            {
                Document doc = uiDoc.Document;

                if (!(uiDoc.ActiveView is View3D activeView))
                {
                    MessageBox.Show(
                        $"Предварительно - открыйте 3д-вид, на котором вы хотите найти элемент",
                        "Внимание",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return Result.Cancelled;
                }

                if (IsDetailLevelControlledByTemplate(doc, activeView, out string templateName))
                {
                    MessageBox.Show(
                        $"Для вида \"{activeView.Name}\" применён шаблон \"{templateName}\", который управляет параметром \"Уровень детализации\".\n\n" +
                        "Сними шаблон с этого вида или настрой шаблон так, чтобы \"Уровень детализации\" мог изменяться вне шаблона.",
                        $"KPLN: {SearchByIdExtCmd.PluginName}: Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);

                    return Result.Cancelled;
                }

                using (Transaction zoomTransaction = new Transaction(doc, "KPLN_Настроить 3D-вид поиска"))
                {
                    zoomTransaction.Start();
                    ZoomElement(uiDoc, activeView);
                    zoomTransaction.Commit();
                }

                RefreshDetailLevel(doc, uiDoc, activeView);

                //Выделение элемента (НЕЛЬЗЯ ДЛЯ <R2023)
#if !Debug2020 && !Revit2020
                Reference reference = new Reference(_searchEnt.Elem);
                Reference linkRef = reference.CreateLinkReference(_searchEnt.ElemDocEntity.SDE_RLI);

                uiDoc.Selection.SetReferences(new List<Reference> { linkRef });
#endif

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Cancelled;
            }
        }

        private static bool IsDetailLevelControlledByTemplate(Document doc, View3D activeView, out string templateName)
        {
            templateName = string.Empty;

            ElementId templateId = activeView.ViewTemplateId;
            if (templateId == null || templateId == ElementId.InvalidElementId)
                return false;

            if (!(doc.GetElement(templateId) is View templateView))
                return false;

            templateName = templateView.Name;
            ICollection<ElementId> nonControlledIds = templateView.GetNonControlledTemplateParameterIds();
            ICollection<ElementId> templateParameterIds = templateView.GetTemplateParameterIds();

            foreach (ElementId parameterId in templateParameterIds)
            {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                bool isDetailLevelParameter = parameterId.IntegerValue == (int)BuiltInParameter.VIEW_DETAIL_LEVEL;
#else
                bool isDetailLevelParameter = parameterId.Value == (long)BuiltInParameter.VIEW_DETAIL_LEVEL;
#endif
                if (!isDetailLevelParameter)
                    continue;

                return !nonControlledIds.Contains(parameterId);
            }

            return false;
        }

        private static void RefreshDetailLevel(Document doc, UIDocument uidoc, View3D activeView)
        {
            try
            {
                using (Transaction coarseTransaction = new Transaction(doc, "KPLN_Детализация 3D-вида: низкая"))
                {
                    coarseTransaction.Start();
                    activeView.DetailLevel = ViewDetailLevel.Coarse;
                    coarseTransaction.Commit();
                }
                uidoc.RefreshActiveView();

                using (Transaction fineTransaction = new Transaction(doc, "KPLN_Детализация 3D-вида: высокая"))
                {
                    fineTransaction.Start();
                    activeView.DetailLevel = ViewDetailLevel.Fine;
                    fineTransaction.Commit();
                }
                uidoc.RefreshActiveView();
            }
            catch (Exception)
            {
                // Если шаблон вида запрещает менять детализацию, продолжаем обычный выбор элемента.
            }
        }

        private void ZoomElement(UIDocument uidoc, View3D activeView)
        {
            BoundingBoxXYZ elemBBox = _searchEnt.GetElemBBox();
            if (elemBBox == null)
                return;

            XYZ offsetMin = new XYZ(0, 0, 0);
            XYZ offsetMax = new XYZ(0, 0, 0);

            activeView.SetSectionBox(new BoundingBoxXYZ() { Max = elemBBox.Max + offsetMax, Min = elemBBox.Min - offsetMin });

            XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
            XYZ up_direction = VectorFromHorizVertAngles(135, 60);
            XYZ centroid = new XYZ((elemBBox.Max.X + elemBBox.Min.X) / 2, (elemBBox.Max.Y + elemBBox.Min.Y) / 2, (elemBBox.Max.Z + elemBBox.Min.Z) / 2);
            ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);

            activeView.SetOrientation(orientation);

            IList<UIView> views = uidoc.GetOpenUIViews();
            foreach (UIView uvView in views)
            {
                if (uvView.ViewId.Equals(activeView.Id))
                    uvView.ZoomAndCenterRectangle(elemBBox.Min, elemBBox.Max);
            }
        }

        private static XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
        {
            double degToRadian = Math.PI * 2 / 360;
            double angleHorizR = angleHorizD * degToRadian;
            double angleVertR = angleVertD * degToRadian;
            double a = Math.Cos(angleVertR);
            double b = Math.Cos(angleHorizR);
            double c = Math.Sin(angleHorizR);
            double d = Math.Sin(angleVertR);

            return new XYZ(a * b, a * c, d);
        }
    }
}

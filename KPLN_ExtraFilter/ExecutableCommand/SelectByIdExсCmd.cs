using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
                Document doc = app.ActiveUIDocument.Document;

                Transaction t = new Transaction(doc, "KPLN_Найти элемент в связи");

                t.Start();

                if (app.ActiveUIDocument.ActiveView is View3D activeView)
                {
                    // Зум к элементу
                    ZoomElement(app.ActiveUIDocument, activeView);


                    //Выделение элемента (НЕЛЬЗЯ ДЛЯ <R2023)
#if !Debug2020 && !Revit2020
                    Reference reference = new Reference(_searchEnt.Elem);
                    Reference linkRef = reference.CreateLinkReference(_searchEnt.ElemDocEntity.SDE_RLI);

                    uiDoc.Selection.SetReferences(new List<Reference> { linkRef });
#endif
                }
                else
                    MessageBox.Show(
                        $"Предварительно - открыйте 3д-вид, на котором вы хотите найти элемент",
                        "Внимание",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                t.Commit();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Cancelled;
            }
        }

        private void ZoomElement(UIDocument uidoc, View3D activeView)
        {
            BoundingBoxXYZ elemBBox = _searchEnt.GetElemBBox();
            if (elemBBox == null)
                return;

            XYZ offsetMin = new XYZ(-5, -5, -2);
            XYZ offsetMax = new XYZ(5, 5, 1);

            activeView.SetSectionBox(new BoundingBoxXYZ() { Max = elemBBox.Max + offsetMax, Min = elemBBox.Min + offsetMin });

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

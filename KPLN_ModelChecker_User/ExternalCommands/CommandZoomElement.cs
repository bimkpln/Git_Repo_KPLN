using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    public class CommandZoomElement : IExecutableCommand
    {
        private Element Element { get; set; }

        private BoundingBoxXYZ Box { get; set; }

        private XYZ Centroid { get; set; }

        private bool _isLegendOpen;

        public CommandZoomElement(Element element)
        {
            Element = element;
        }

        public CommandZoomElement(Element element, BoundingBoxXYZ box, XYZ centroid) : this(element)
        {
            Box = box;
            Centroid = centroid;
        }

        public Result Execute(UIApplication app)
        {
            try
            {
                CutView(Box, Centroid, app.ActiveUIDocument, Element);
                ICollection<ElementId> newSelection = new List<ElementId>() { Element.Id };
                app.ActiveUIDocument.Selection.SetElementIds(newSelection);
            }
            catch (Exception)
            {
                _isLegendOpen = false;

                try
                {
                    // Легенды Ревит не умеет подбирать. Добавлен вывод на экран сообщения, чтобы открыли вид вручную
                    app.DialogBoxShowing += new EventHandler<DialogBoxShowingEventArgs>(DialogBoxShowing);
                    
                    app.ActiveUIDocument.ShowElements(Element);
                    if (_isLegendOpen)
                    {
                        Dimension dim = Element as Dimension;
                        View view = dim.View;
                        TaskDialog.Show("KPLN", $"Открой легенду ({view.Name}) вручную");
                    }
                    app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>() { Element.Id });
                    
                    app.DialogBoxShowing -= new EventHandler<DialogBoxShowingEventArgs>(DialogBoxShowing);
                }
                catch (Exception) { }
            }
            return Result.Succeeded;
        }
        private void CutView(BoundingBoxXYZ box, XYZ centroid, UIDocument uidoc, Element element)
        {
            XYZ offsetMin = new XYZ(-5, -5, -2);
            XYZ offsetMax = new XYZ(5, 5, 1);
            View3D activeView = uidoc.ActiveView as View3D;
            ViewFamily activeViewFamily = ViewFamily.Invalid;
            ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
            activeViewFamily = viewFamilyType.ViewFamily;
            if (activeViewFamily == ViewFamily.ThreeDimensional)
            {
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Min = box.Min + new XYZ(-1, -1, -5);
                sectionBox.Max = box.Max + new XYZ(1, 1, 5);
                activeView.SetSectionBox(sectionBox);
                XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);
                activeView.SetOrientation(orientation);
                IList<UIView> views = uidoc.GetOpenUIViews();
                foreach (UIView uvView in views)
                {
                    if (uvView.ViewId.IntegerValue == activeView.Id.IntegerValue)
                    {
                        uvView.ZoomAndCenterRectangle(box.Min, box.Max);
                        break;
                    }
                }
            }
            else
            {
                uidoc.ShowElements(element);
            }
        }
        private XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
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

        /// <summary>
        /// Закрывает окно с ошибкой об открытии легенды. Меняет параметр _isLegendOpen на true
        /// </summary>
        private void DialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            TaskDialogShowingEventArgs td = args as TaskDialogShowingEventArgs;
            if (td.Message.Equals("Невозможно подобрать подходящий вид."))
            {
                args.OverrideResult(1);
                _isLegendOpen = true;
            }
        }
    }
}

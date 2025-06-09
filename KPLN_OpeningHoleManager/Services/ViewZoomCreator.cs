using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Controls;
using System.Xml.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Класс-помощник в создании спец. вида в Revit
    /// </summary>
    internal static class ViewZoomCreator
    {
        private static string _viewForObserveName;

        internal static string ViewForObserveName
        {
            get
            {
                if (_viewForObserveName == null)
                    _viewForObserveName = $"OHM_В_{MainDBService.CurrentDBUser.RevitUserName}";

                return _viewForObserveName;
            }
        }

        /// <summary>
        /// Метод для создания нужного вида
        /// </summary>
        internal static View3D SpecialViewCreator(UIApplication uiapp, IEnumerable<Element> elemsToObserve, bool ignoreFullyVisible)
        {
            if (ignoreFullyVisible)
                return null;
            
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            BoundingBoxXYZ hostElemsBbox = SumBBox(elemsToObserve);
            
            View activeView = uidoc.ActiveView;
            // Если уже открыт 3д-вид, или разрез - то можно игнорировать процесс создания нового вида
            if (activeView is View3D || activeView is ViewSection)
            {
                
                foreach(Element elem in elemsToObserve)
                {
                    if (IsElementFullyVisibleOnView(elem, activeView))
                        return null;
                }
            }

            View viewForUser = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .WhereElementIsNotElementType()
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .FirstOrDefault(v => v.Name.Contains(ViewForObserveName));

            if (viewForUser == null)
            {
                var view_type = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .Where(vft => vft.ViewFamily == ViewFamily.ThreeDimensional)
                    .FirstOrDefault();

                View3D newView = View3D.CreateIsometric(doc, view_type.Id);
                newView.Name = ViewForObserveName;
                SetViewSettings(doc, newView);

                viewForUser = newView;

                doc.Regenerate();
            }

            View3D viewForUser3D = viewForUser as View3D;
            ZoomElement(uidoc, hostElemsBbox, viewForUser3D);

            return viewForUser3D;
        }

        /// <summary>
        /// Метод для открытия нужного вида
        /// </summary>
        internal static void SpecialViewOpener(UIApplication uiapp, View view)
        {
            if (view != null)
                uiapp.ActiveUIDocument.RequestViewChange(view);
        }

        /// <summary>
        /// Проверка элемента на ПОЛНОЕ наличие на виде
        /// </summary>
        private static bool IsElementFullyVisibleOnView(Element element, View view)
        {
            // Атрыманне BoundingBox элемента на гэтым выглядзе
            BoundingBoxXYZ elementBB = element.get_BoundingBox(view);
            if (elementBB == null)
                return false;

            // Атрыманне CropBox выгляду
            BoundingBoxXYZ cropBox = view.CropBox;
            if (cropBox == null || !view.CropBoxActive)
                return true;

            // Атрыманне трансфармацыі выгляду
            Transform viewTransform = cropBox.Transform;

            // Праверка, ці ўсе 8 кропак BoundingBox элемента ў CropBox
            List<XYZ> elementPoints = GetBoundingBoxCorners(elementBB);

            foreach (var pt in elementPoints)
            {
                // Пераход у лакальныя каардынаты CropBox
                XYZ localPt = viewTransform.Inverse.OfPoint(pt);

                if (!IsPointInsideBox(localPt, cropBox.Min, cropBox.Max))
                {
                    return false; // хаця б адна кропка па-за межамі
                }
            }

            return true;
        }

        private static List<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ bb)
        {
            // Стварэнне спісу з 8 кропак BoundingBox
            var min = bb.Min;
            var max = bb.Max;

            return new List<XYZ>
            {
                new XYZ(min.X, min.Y, min.Z),
                new XYZ(min.X, min.Y, max.Z),
                new XYZ(min.X, max.Y, min.Z),
                new XYZ(min.X, max.Y, max.Z),
                new XYZ(max.X, min.Y, min.Z),
                new XYZ(max.X, min.Y, max.Z),
                new XYZ(max.X, max.Y, min.Z),
                new XYZ(max.X, max.Y, max.Z)
            }.Select(p => bb.Transform.OfPoint(p)).ToList();
        }

        private static bool IsPointInsideBox(XYZ point, XYZ min, XYZ max)
        {
            return point.X >= min.X && point.X <= max.X &&
                   point.Y >= min.Y && point.Y <= max.Y &&
                   point.Z >= min.Z && point.Z <= max.Z;
        }

        private static void SetViewSettings(Document doc, View3D view3D)
        {
            #region Дисциплина и уровень детализации
            view3D.Discipline = ViewDiscipline.Mechanical;
            view3D.DetailLevel = ViewDetailLevel.Fine;
            #endregion

            #region Настройка фильтров по имени типа
            string filterName = $"OHM_Стена = !*{string.Join(" ИЛИ !*", ARKRElemsWorker.ARKRNames_StartWith)}";
            
            ParameterFilterElement[] docFilters = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .ToArray();


            List<ElementFilter> efColl = new List<ElementFilter>();
            for (int i = 0; i < ARKRElemsWorker.ARKRNames_StartWith.Count(); i++)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId((int)BuiltInParameter.SYMBOL_NAME_PARAM), ARKRElemsWorker.ARKRNames_StartWith[i], false);
                ElementParameterFilter elementParameterFilter = new ElementParameterFilter(fRule);
                efColl.Add(elementParameterFilter);
            }

            LogicalAndFilter finalFRule = new LogicalAndFilter(efColl);


            ElementId viewFilterId = null;
            ParameterFilterElement oldEqualFRule = docFilters.FirstOrDefault(df => df.Name.Equals(filterName));
            if (oldEqualFRule != null)
                viewFilterId = oldEqualFRule.Id;
            else
            {
                List<ElementId> catId = new List<ElementId>() { new ElementId((int)BuiltInCategory.OST_Walls) };
                ParameterFilterElement newViewFilter = ParameterFilterElement.Create(doc, filterName, catId, finalFRule);
                viewFilterId = newViewFilter.Id;
            }

            view3D.AddFilter(viewFilterId);
            view3D.SetFilterVisibility(viewFilterId, false);
            #endregion
        }

        /// <summary>
        /// Создать общий BoundingBoxXYZ для элементов
        /// </summary>
        private static BoundingBoxXYZ SumBBox(IEnumerable<Element> elems)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (Element element in elems)
            {
                BoundingBoxXYZ elementBox = element.get_BoundingBox(null);
                if (elementBox == null)
                    continue;

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = elementBox.Min,
                        Max = elementBox.Max
                    };
                }
                else
                {
                    resultBBox.Min = new XYZ(
                        Math.Min(resultBBox.Min.X, elementBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elementBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elementBox.Min.Z));

                    resultBBox.Max = new XYZ(
                        Math.Max(resultBBox.Max.X, elementBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elementBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elementBox.Max.Z));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Зум по указанному BoundingBoxXYZ
        /// </summary>
        private static void ZoomElement(UIDocument uidoc, BoundingBoxXYZ box, View3D viewToZoom)
        {
            XYZ offsetMin = new XYZ(-1, -1, 0);
            XYZ offsetMax = new XYZ(1, 1, 0);

            ViewFamily activeViewFamily = ViewFamily.Invalid;
            try
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(viewToZoom.GetTypeId()) as ViewFamilyType;
                activeViewFamily = viewFamilyType.ViewFamily;
            }
            catch (Exception) { }
            if (activeViewFamily == ViewFamily.ThreeDimensional)
            {
                viewToZoom.SetSectionBox(new BoundingBoxXYZ() { Max = box.Max + offsetMax, Min = box.Min + offsetMin });
                XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                XYZ centroid = new XYZ((box.Max.X + box.Min.X) / 2, (box.Max.Y + box.Min.Y) / 2, (box.Max.Z + box.Min.Z) / 2);
                ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);
                viewToZoom.SetOrientation(orientation);
                IList<UIView> views = uidoc.GetOpenUIViews();
                foreach (UIView uvView in views)
                {
                    if (uvView.ViewId.IntegerValue == viewToZoom.Id.IntegerValue)
                    {
                        uvView.ZoomAndCenterRectangle(box.Min, box.Max);
                    }
                }
                return;
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

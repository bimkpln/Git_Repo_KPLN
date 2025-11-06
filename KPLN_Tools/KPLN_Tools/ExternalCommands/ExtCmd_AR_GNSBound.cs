using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    /// <summary>
    /// Метод с экстрапомещениям взят отсюда: https://thebuildingcoder.typepad.com/blog/2018/05/filterrule-use-and-retrieving-exterior-walls.html 
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_AR_GNSBound : IExternalCommand
    {
        private ModelCurveArray _exteriorModelCurveArr;
        private Room _exteriorRoom;


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Добавить проверку что это план и т.п., 
            ViewPlan activeViewPlan;
            View activeView = uidoc.ActiveView;
            if (activeView is ViewPlan vPlan)
                activeViewPlan = vPlan;
            else
            {
                TaskDialog.Show(
                    "Ошибка",
                    "Работа остановлена. Работает ТОЛЬКО на открытом плане этажа",
                    TaskDialogCommonButtons.Cancel);

                return Result.Cancelled;
            }

            // Текущий уровень плана
            Level planGenLvl = activeViewPlan.GenLevel;

            using (Transaction transRoom = new Transaction(doc, "KPLN_Контур_Помещение"))
            {
                transRoom.Start();

                // Получаю границы помещения
                XYZ[] pntForRoomBorder = GetPointsForRoomBorder(doc, activeView, planGenLvl);
                if (!pntForRoomBorder.Any())
                {
                    TaskDialog.Show(
                        "Ошибка",
                        "Работа остановлена. Не удалось получить точки для рамки внешнего габарита технического помещения. Отправь разработчику",
                        TaskDialogCommonButtons.Cancel);

                    return Result.Cancelled;
                }

                // Создаю линни границы
                if (!CreateExteriorLines(doc, planGenLvl, pntForRoomBorder))
                    return Result.Cancelled;

                // Создаю помещение
                CreateExteriorRoom(doc, planGenLvl, pntForRoomBorder);
                if (_exteriorRoom == null)
                {
                    TaskDialog.Show(
                        "Ошибка",
                        "Работа остановлена. Не удалось создать техническое помещение. Отправь разработчику",
                        TaskDialogCommonButtons.Cancel);

                    return Result.Cancelled;
                }

                transRoom.Commit();
            }

            // Добавить выбор типа "Область маскировки"
            FilledRegionType frt = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            // Анализирую внешние стены - получаю линии
            List<CurveLoop> insideBoundCurveLoops = GetInsideRoomCurveLoops(doc, _exteriorRoom);
            using (Transaction transCreate = new Transaction(doc, "KPLN_Контур_Создание и чистка"))
            {
                transCreate.Start();

                if (insideBoundCurveLoops.Count() == 0)
                {
                    TaskDialog.Show(
                        "Ошибка",
                        "Работа откачена. Не удалось получть контур внешних стен. Если их НЕТ на текущем виде, то нужно добавить. В противном случае - отправь разработчику",
                        TaskDialogCommonButtons.Cancel);
                }
                else
                {
                    // Создаю зону (НЕ ОЧЕНЬ ВАРИК, Т.К. НЕ РАБОТАЕТ С ВАРИАНТАМИ - ЗОНЫ И ПРОСТРАНСТВА ТАМ НЕ СОЗДАЮТСЯ)
                    //SketchPlane sketchPlane = SketchPlane.Create(doc, planGenLvl.Id);
                    //foreach(CurveLoop cLoop in insideBoundCurveLoops)
                    //{
                    //    foreach(Curve curve in cLoop)
                    //    {
                    //        ModelCurve boudLine = doc.Create.NewAreaBoundaryLine(sketchPlane, curve, activeViewPlan);
                    //    }
                    //}

                    // Создаю заливку
                    FilledRegion.Create(doc, frt.Id, activeView.Id, insideBoundCurveLoops);
                }

                DocClearing(doc);

                transCreate.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию точек для границы помещения
        /// </summary>
        private static XYZ[] GetPointsForRoomBorder(Document doc, View currentView, Level downPrjLvl)
        {
            IEnumerable<BoundingBoxXYZ> bboxColl = new FilteredElementCollector(doc, currentView.Id)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Select(el => el.get_BoundingBox(null));

            // Получаю мин и макс экстра точки
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            foreach (BoundingBoxXYZ windBox in bboxColl)
            {
                XYZ min = windBox.Min;
                if (min.X < minX) minX = min.X;
                if (min.Y < minY) minY = min.Y;

                XYZ max = windBox.Max;
                if (max.X > maxX) maxX = max.X;
                if (max.Y > maxY) maxY = max.Y;
            }
            double minZ = downPrjLvl.Elevation;
            XYZ minExtraPoint = new XYZ(minX, minY, minZ);
            XYZ maxExtraPoint = new XYZ(maxX, maxY, minZ);

            // Расширяю ганицы
            XYZ expandXYZ = new XYZ(20, 20, 0);
            XYZ resultMinPnt = minExtraPoint + expandXYZ.Negate();
            XYZ resultMaxPnt = maxExtraPoint + expandXYZ;

            // Выставляю рамку и приземляю его в минимум по Z
            XYZ[] resultPoints = new XYZ[]
            {
                resultMinPnt,
                new XYZ(resultMaxPnt.X, resultMinPnt.Y, minZ),
                new XYZ(resultMaxPnt.X, resultMaxPnt.Y, minZ),
                new XYZ(resultMinPnt.X, resultMaxPnt.Y, minZ),
            };

            return resultPoints;
        }

        /// <summary>
        /// Получить смещенную точку для помещения (смещение нужно чтобы помещения не накладывалист)
        /// </summary>
        private static XYZ GetExcenterPnt(XYZ[] points)
        {
            XYZ lastPnt = points[points.Length - 1];
            XYZ tempPnt = lastPnt;

            XYZ mutablePnt = tempPnt / 1.05;
            return new XYZ(mutablePnt.X, mutablePnt.Y, lastPnt.Z);
        }

        /// <summary>
        /// Получить коллекцию кривых CurveLoop, которые являются внутренними границами экстрапомещения (т.е. наружные стены здания)
        /// </summary>
        /// <returns></returns>
        private static List<CurveLoop> GetInsideRoomCurveLoops(Document doc, Room room)
        {
            List<CurveLoop> result = new List<CurveLoop>();

            IList<IList<BoundarySegment>> bsCollection = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

            foreach (IList<BoundarySegment> currentBSList in bsCollection)
            {
                // Игнорирую границы помещений
                if (currentBSList.All(bs => doc.GetElement(bs.ElementId) is ModelLine))
                    continue;

                List<XYZ> allBoundPoint = new List<XYZ>();
                foreach (BoundarySegment bs in currentBSList)
                {
                    Element boundElem = doc.GetElement(bs.ElementId);

                    Curve boundCurve = null;
                    // Анализирую элементы из связи
                    if (boundElem is RevitLinkInstance rli)
                    {
                        // add link elem logic, first of all - coord transf
                    }
                    // Анализирую элементы из модели
                    else
                        boundCurve = bs.GetCurve();

                    if (boundCurve != null)
                    {
                        allBoundPoint.Add(boundCurve.GetEndPoint(0));
                        allBoundPoint.Add(boundCurve.GetEndPoint(1));
                    }
                }

                List<XYZ> simplifyPnts = SimplifyCurves(allBoundPoint, 0.1);
                if (simplifyPnts.Count > 3)
                {
                    List<Curve> curves = new List<Curve>();
                    for (int i = 0; i < simplifyPnts.Count; i++)
                    {
                        XYZ start = simplifyPnts[i];
                        XYZ end = simplifyPnts[(i + 1) % simplifyPnts.Count]; // Замыкаем контур

                        if (!start.IsAlmostEqualTo(end, 0.0001))
                        {
                            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(start, end);
                            curves.Add(line);
                        }
                    }

                    CurveLoop curveLoopUp = CurveLoop.Create(curves);
                    if (!curveLoopUp.IsOpen())
                        result.Add(curveLoopUp);
                }
            }

            return result;
        }

        public static List<XYZ> SimplifyCurves(List<XYZ> points, double tolerance)
        {
            if (points.Count < 3)
                return new List<XYZ>(points);

            List<XYZ> simplified = new List<XYZ> { points[0] };

            for (int i = 1; i < points.Count - 1; i++)
            {
                XYZ prev = simplified.Last();
                XYZ current = points[i];
                XYZ next = points[i + 1];

                if (!IsCollinear(prev, current, next, tolerance))
                {
                    simplified.Add(current);
                }
            }

            simplified.Add(points.Last());
            return simplified;
        }

        private static bool IsCollinear(XYZ a, XYZ b, XYZ c, double tolerance)
        {
            XYZ ab = b - a;
            XYZ bc = c - b;
            return ab.CrossProduct(bc).GetLength() < tolerance;
        }

        /// <summary>
        /// Собираю максимальный внешний контур здания по окнам
        /// </summary>
        private bool CreateExteriorLines(Document doc, Level downPrjLvl, XYZ[] pntForRoomBorder)
        {
            try
            {
                SketchPlane sketchPlane = SketchPlane.Create(doc, downPrjLvl.Id);
                CurveArray curveArray = new CurveArray();
                for (int i = 0; i < pntForRoomBorder.Count(); i++)
                {
                    Autodesk.Revit.DB.Line line = null;
                    if (i == pntForRoomBorder.Count() - 1)
                        line = Autodesk.Revit.DB.Line.CreateBound(pntForRoomBorder[i], pntForRoomBorder[0]);
                    else
                        line = Autodesk.Revit.DB.Line.CreateBound(pntForRoomBorder[i], pntForRoomBorder[i + 1]);

                    curveArray.Append(line);
                }

                // Рандомный вид для создания границ (они всё равно будут удалены)
                ViewPlan randomDocPlan = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .WhereElementIsNotElementType()
                    .Cast<ViewPlan>()
                    .Where(vp => !vp.IsTemplate)
                    .FirstOrDefault();

                _exteriorModelCurveArr = doc.Create.NewRoomBoundaryLines(sketchPlane, curveArray, randomDocPlan);

                return true;
            }
            catch (Exception ex)
            {
                HtmlOutput.Print($"Работа остановлена. Не удалось создать рамку для внешнего габарита технического помещения. Отправь разработчику: {ex.Message}", MessageType.Error);
            }

            return false;
        }

        /// <summary>
        /// Создаю помещение по внешнему контуру
        /// </summary>
        private void CreateExteriorRoom(Document doc, Level downPrjLvl, XYZ[] pntForRoomBorder)
        {
            XYZ roomLocation = GetExcenterPnt(pntForRoomBorder);
            _exteriorRoom = doc.Create.NewRoom(downPrjLvl, new UV(roomLocation.X, roomLocation.Y));
            _exteriorRoom.get_Parameter(BuiltInParameter.ROOM_NAME).Set("BIM: Создание рамки для ГНС. Если читаешь - УДАЛИ это помещение, оно для технических нужд!");
        }

        /// <summary>
        /// Чистка от помещения и границ
        /// </summary>
        private void DocClearing(Document doc)
        {
            doc.Delete(_exteriorRoom.Id);
            foreach (var obj in _exteriorModelCurveArr)
            {
                ModelLine modelLine = obj as ModelLine;
                doc.Delete(modelLine.Id);
            }
        }
    }
}

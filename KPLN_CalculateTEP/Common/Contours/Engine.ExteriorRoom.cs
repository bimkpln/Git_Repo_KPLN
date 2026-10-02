using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private Solid ExteriorRoomContour(List<Record> walls, Record floor, Indicator indicator, double elevation)
            {
                // An outside room sees Finish/CoreBoundary from outside the building. The inside
                // face used for gross areas still requires the existing wall-section algorithm.
                if ((int)indicator > 4) throw new InvalidOperationException("Для внутренней границы этажа используется сечение стен.");
                var lastPhase = doc.Phases.Cast<Phase>().Last();
                if (!string.IsNullOrWhiteSpace(Config.Phase) && !Eq(lastPhase.Name, Config.Phase))
                    throw new InvalidOperationException("Временное внешнее помещение доступно для последней стадии основной модели.");
                foreach (var source in walls.Select(w => w.Source).Distinct())
                {
                    RequireHorizontalTransform(source.Transform);
                    if (source.Document.Equals(doc)) continue;
                    var link = source.RootLink == null ? null : doc.GetElement(source.RootLink) as RevitLinkInstance;
                    if (link == null || !link.GetLinkDocument().Equals(source.Document))
                        throw new InvalidOperationException("Вложенная связь требует расчёта по сечениям стен: " + source.Name);
                    var bounding = doc.GetElement(link.GetTypeId()).get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                    if (bounding == null || bounding.AsInteger() != 1)
                        throw new InvalidOperationException("Связь не ограничивает помещения: " + source.Name);
                    if (!Eq(source.Document.Phases.Cast<Phase>().Last().Name, lastPhase.Name))
                        throw new InvalidOperationException("Стадии основной модели и связи требуют расчёта по сечениям: " + source.Name);
                    // Custom phase mappings must not silently select different linked geometry.
                    var phaseMap = ((RevitLinkType)doc.GetElement(link.GetTypeId())).GetPhaseMap();
                    ElementId mapped;
                    if (!phaseMap.TryGetValue(lastPhase.Id, out mapped) || mapped != source.Document.Phases.Cast<Phase>().Last().Id)
                        throw new InvalidOperationException("Связь сопоставляет другую стадию: " + source.Name);
                }
                var points = new List<XYZ>();
                // Enclose every potentially room-bounding wall, including interior-labelled walls,
                // so the frame can never accidentally cut through an unselected part of a building.
                foreach (var source in Sources.Where(s => s.Loaded && s.LoadError == null))
                    foreach (var wall in source.Elements.OfType<Wall>().Where(w => PhaseAccepted(source, w) && CrossesFloor(w, source.Transform, elevation)))
                    {
                        var box = wall.get_BoundingBox(null); if (box == null) continue;
                        var t = source.Transform.Multiply(box.Transform);
                        for (int i = 0; i < 8; i++) points.Add(t.OfPoint(new XYZ((i & 1) == 0 ? box.Min.X : box.Max.X, (i & 2) == 0 ? box.Min.Y : box.Max.Y, (i & 4) == 0 ? box.Min.Z : box.Max.Z)));
                    }
                if (points.Count == 0) throw new InvalidOperationException("Не найден габарит стен для внешнего помещения.");
                const double margin = 20; // XY working frame only; never a measurement-height offset.
                double left = points.Min(p => p.X) - margin, right = points.Max(p => p.X) + margin;
                double bottom = points.Min(p => p.Y) - margin, top = points.Max(p => p.Y) + margin;
                var corners = new[] { new XYZ(left, bottom, elevation), new XYZ(right, bottom, elevation), new XYZ(right, top, elevation), new XYZ(left, top, elevation) };
                var copied = new List<CurveLoop>();
                using (var transaction = new Transaction(doc, "ТЭП: временное внешнее помещение"))
                {
                    transaction.Start();
                    try
                    {
                        var options = transaction.GetFailureHandlingOptions().SetClearAfterRollback(true);
                        transaction.SetFailureHandlingOptions(options);
                        var level = Level.Create(doc, elevation);
                        var height = level.get_Parameter(BuiltInParameter.LEVEL_ROOM_COMPUTATION_HEIGHT);
                        if (height == null || height.IsReadOnly || !height.Set(0)) throw new InvalidOperationException("Нельзя установить вычисление границ на отметке пола.");
                        var type = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(t => t.ViewFamily == ViewFamily.FloorPlan);
                        if (type == null) throw new InvalidOperationException("В модели отсутствует тип плана этажа.");
                        var plan = ViewPlan.Create(doc, type.Id, level.Id);
                        var sketch = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, corners[0]));
                        var curves = new CurveArray();
                        for (int i = 0; i < 4; i++) curves.Append(Line.CreateBound(corners[i], corners[(i + 1) % 4]));
                        var lines = doc.Create.NewRoomBoundaryLines(sketch, curves, plan);
                        var frameIds = new HashSet<ElementId>(lines.Cast<ModelCurve>().Select(l => l.Id));
                        // Point is chosen relative to the frame, independent of model origin/sign.
                        var room = doc.Create.NewRoom(level, new UV(left + margin / 2, bottom + margin / 2));
                        room.BaseOffset = 0;
                        doc.Regenerate();
                        if (room.Area <= 0 || Math.Abs(height.AsDouble()) > FloorGeometryTolerance || Math.Abs(level.ProjectElevation - elevation) > FloorGeometryTolerance)
                            throw new InvalidOperationException("Внешнее помещение не замкнуто на заданной отметке пола.");
                        var boundaryOptions = new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = Methodology.UsesCoreBoundary ? SpatialElementBoundaryLocation.CoreBoundary : SpatialElementBoundaryLocation.Finish };
                        var boundaries = room.GetBoundarySegments(boundaryOptions);
                        if (boundaries == null) throw new InvalidOperationException("Revit не вернул границы внешнего помещения.");
                        var seen = new HashSet<string>();
                        foreach (var ring in boundaries)
                        {
                            int frame = ring.Count(s => frameIds.Contains(s.ElementId));
                            if (frame != 0)
                            {
                                if (frame != ring.Count) throw new InvalidOperationException("Внешнее помещение проникло в здание: наружный контур разомкнут.");
                                continue;
                            }
                            foreach (var segment in ring)
                            {
                                var element = doc.GetElement(segment.ElementId);
                                string key = "host/" + IDHelper.ElIdValue(segment.ElementId);
                                if (element is RevitLinkInstance)
                                {
                                    var link = (RevitLinkInstance)element;
                                    element = link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                                    key = IDHelper.ElIdValue(link.Id) + "/" + IDHelper.ElIdValue(segment.LinkElementId);
                                }
                                var wall = element as Wall;
                                if (wall == null || wall.WallType.Function != WallFunction.Exterior)
                                    throw new InvalidOperationException("Контур внешнего помещения содержит разделитель или элемент, не обозначенный наружной стеной. Нужна проверка оболочки.");
                                seen.Add(key);
                            }
                            // Preserve actual arcs, ellipses and splines. Never join just endpoints.
                            var loop = CurveLoop.Create(ring.Select(s => s.GetCurve().Clone()).ToList());
                            if (loop.IsOpen()) throw new InvalidOperationException("Revit вернул незамкнутую петлю внешнего помещения.");
                            copied.Add(loop);
                        }
                        if (copied.Count == 0) throw new InvalidOperationException("Внутри внешнего помещения нет замкнутой оболочки здания.");
                        // An unseen external wall may enclose a courtyard. An outside probe cannot
                        // see a disconnected courtyard, so never silently fill it into the GNS.
                        if (walls.Any(w => !seen.Contains((w.Source.Document.Equals(doc) ? "host" : IDHelper.ElIdValue(w.Source.RootLink).ToString()) + "/" + IDHelper.ElIdValue(w.Element.Id))))
                            throw new InvalidOperationException("Есть наружные стены, недоступные внешнему помещению (возможен внутренний двор или дополнительная оболочка).");
                    }
                    finally
                    {
                        if (transaction.GetStatus() == TransactionStatus.Started && transaction.RollBack() != TransactionStatus.RolledBack)
                            throw new InvalidOperationException("Не удалось откатить временное внешнее помещение.");
                        DisposeSpatialCalculators(); localSpatialVolumes.Clear(); parameterCache.Clear(); floorWallFaces.Clear(); floorFaceHeights.Clear();
                    }
                }
                // Detached copies survive rollback. All temporary model objects are already gone.
                var flat = copied.Select(l => FlatLoop(l)).ToList();
                // Validate area, not merely visual chord deviation. Each disconnected contour is
                // compared separately; opposite approximation errors cannot cancel between buildings.
                double tolerance = ArcChordTolerance;
                var exactAreas = flat.Select(l => Extrude(new List<CurveLoop> { l }).Volume).ToList();
                Solid result = null;
                for (int attempt = 0; attempt < 6; attempt++, tolerance /= 4)
                {
                    bool accurate = true;
                    for (int i = 0; i < flat.Count; i++)
                    {
                        var part = PlanFromSectionLoops(new[] { flat[i] }, tolerance);
                        if (part == null || Math.Abs(part.Volume - exactAreas[i]) * .09290304 > .001) { accurate = false; break; }
                    }
                    if (!accurate) continue;
                    result = PlanFromSectionLoops(flat, tolerance); break;
                }
                if (result == null) throw new InvalidOperationException("Точность площади внешнего контура 0,001 м² относительно исходных кривых Revit не подтверждена.");
                Notice("EXTERIOR_ROOM", "Информация", "Контур получен через временное внешнее помещение на отметке пола. Площадь каждой петли сверена с исходными кривыми Revit (расхождение до 0,001 м² до вычитания исключений). Временные помещение, рамка, уровень и вид отменены транзакцией.", floor, indicator.ToString());
                return result;
            }
        }
    }
}

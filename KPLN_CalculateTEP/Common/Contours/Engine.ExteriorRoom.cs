using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private class UnconfirmedEnvelopeException : InvalidOperationException
            {
                internal PlanarRegion Candidate {get;}
                internal string CandidateDescription {get;}
                internal UnconfirmedEnvelopeException(string message,PlanarRegion candidate=null,string candidateDescription=null) : base(message)
                {Candidate=candidate;CandidateDescription=candidateDescription;}
            }

            private sealed class ExteriorFrameConflictException : UnconfirmedEnvelopeException
            {
                internal ExteriorFrameConflictException(string message,PlanarRegion candidate,string description)
                    :base(message,candidate,description){}
            }

            internal static void ValidateExteriorCoverage(PlanarRegion envelope,PlanarRegion occupied,PlanarRegion material,bool allowUnclassified=false)
            {
                if(envelope==null||envelope.IsEmpty)
                    throw new UnconfirmedEnvelopeException("Не получена замкнутая оболочка этажа.");
                double missing=RegionDifference(occupied,envelope).Area*.09290304;
                if(missing>.005)
                    throw new UnconfirmedEnvelopeException("Вне построенного контура осталось "+missing.ToString("0.###",CultureInfo.InvariantCulture)+
                        " м² помещений или шахт. Граница этажа неполная; проверьте ограждения и разделители помещений.",envelope,
                        "Построена только часть контура этажа. Предварительная область сохранена для исправления, в итог ТЭП не включена.");
                double unknown=RegionDifference(envelope,RegionUnion(occupied,material)).Area*.09290304;
                if(unknown>.005&&!allowUnclassified)
                    throw new UnconfirmedEnvelopeException("Внутри оболочки не классифицировано "+unknown.ToString("0.###",CultureInfo.InvariantCulture)+
                        " м²: область не занята помещениями или сечениями ограждений. Возможен двор, проём или участок без помещений; назначение не подменено автоматически.",envelope,
                        "Контур построен, но внутри есть неподтверждённая площадь "+unknown.ToString("0.###",CultureInfo.InvariantCulture)+
                        " м². Область сохранена для проверки и исправления, в итог ТЭП не включена.");
            }

            private PlanarRegion RecoverExteriorFrameRegion(List<Record> walls,Record floor,Indicator indicator,double elevation,List<Record> rooms,ExteriorFrameConflictException original)
            {
                PlanarRegion candidate=null;
                try
                {
                    var metric=Config.Metrics.First(m=>m.Key==indicator.ToString());
                    var exterior=walls.Where(r=>ContourWallSelected(r,metric)).ToList();
                    if(exterior.Count==0)throw new InvalidOperationException("На расчётной отметке не найдены выбранные наружные стены.");
                    candidate=WallContourRegion(exterior,metric,indicator,elevation);
                    var occupied=rooms.Aggregate(AutomaticShaftFloorRegion(elevation),(r,item)=>RegionUnion(r,RoomNetRegion(item)));
                    var material=FloorMaterialRegion(PhysicalShellCandidates(walls,floor,elevation),elevation,indicator,candidate);
                    ValidateExteriorCoverage(candidate,occupied,material);
                    Notice("EXTERIOR_FRAME_WALL_RECOVERY","Информация","Внешнее помещение не дало отдельной границы здания. Контур восстановлен по геометрии наружных стен на той же отметке; проверены замкнутость и охват помещений.",floor,indicator.ToString());
                    return candidate;
                }
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex)
                {
                    var recovered=ex as UnconfirmedEnvelopeException;
                    bool useRecovered=recovered?.Candidate!=null&&!recovered.Candidate.IsEmpty;
                    bool hasWallCandidate=candidate!=null&&!candidate.IsEmpty;
                    throw new UnconfirmedEnvelopeException(original.Message+" Проверка по наружным стенам: "+ex.Message,
                        useRecovered?recovered.Candidate:hasWallCandidate?candidate:original.Candidate,
                        useRecovered?recovered.CandidateDescription:hasWallCandidate?
                            "Контур построен по геометрии наружных стен, но его полнота не подтверждена. Предварительная область сохранена для проверки, в итог ТЭП не включена.":original.CandidateDescription);
                }
            }

            private PlanarRegion ExteriorRoomRegion(List<Record> walls, Record floor, Indicator indicator, double elevation, List<Record> rooms)
            {
                bool interior = (int)indicator > 4;
                var shaftAreas=AutomaticShaftFloorRegion(elevation);
                var physicalCandidates=PhysicalShellCandidates(walls,floor,elevation);
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
                var points = ExteriorFramePoints(walls,physicalCandidates,elevation);
                if (points.Count == 0) throw new InvalidOperationException("Не найден габарит стен для внешнего помещения.");
                const double margin = 20; // XY working frame only; never a measurement-height offset.
                double left = points.Min(p => p.X) - margin, right = points.Max(p => p.X) + margin;
                double bottom = points.Min(p => p.Y) - margin, top = points.Max(p => p.Y) + margin;
                var corners = new[] { new XYZ(left, bottom, elevation), new XYZ(right, bottom, elevation), new XYZ(right, top, elevation), new XYZ(left, top, elevation) };
                var nativeRegions = new List<PlanarRegion>();
                var innerBands=PlanarRegion.Empty;
                var interiorFailures=new List<string>();
                var exteriorFailures=new List<string>();
                var frameFailures=new List<string>();
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
                        var boundaryOptions = new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = !interior && Methodology.UsesCoreBoundary ? SpatialElementBoundaryLocation.CoreBoundary : SpatialElementBoundaryLocation.Finish };
                        var boundaries = room.GetBoundarySegments(boundaryOptions);
                        if (boundaries == null) throw new InvalidOperationException("Revit не вернул границы внешнего помещения.");
                        var occupied=rooms.Aggregate(shaftAreas,(r,item)=>RegionUnion(r,RoomNetRegion(item)));

                        foreach (var ring in boundaries)
                        {
                            var frame = ring.Select(segment => {
                                if(frameIds.Contains(segment.ElementId))return true;
                                var line=segment.GetCurve() as Line;
                                if(line==null)return false;
                                var a=line.GetEndPoint(0);var b=line.GetEndPoint(1);
                                return Math.Abs(a.Z-elevation)<=FloorGeometryTolerance&&Math.Abs(b.Z-elevation)<=FloorGeometryTolerance&&
                                    IsCalculationFrameEdge(a.X,a.Y,b.X,b.Y,left,bottom,right,top,FloorGeometryTolerance);
                            }).ToArray();
                            if(frame.All(value=>value))continue;
                            bool mixedFrame=frame.Any(value=>value);
                            var segments=ring.Where((s,i)=>!frame[i]).ToList();
                            var nativeCurves=segments.Select(s=>s.GetCurve().Clone()).ToList();
                            PlanarRegion native;
                            try
                            {
                                native=NativeBoundaryRegion(ring.Select(s=>s.GetCurve().Clone()).ToList(),elevation,frame);
                            }
                            catch(OperationCanceledException){throw;}
                            catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                            catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                            catch(Exception ex)
                            {
                                if(mixedFrame)frameFailures.Add(DescribeExteriorFrameConflict(ring,frame,elevation)+" "+ex.Message);
                                else exteriorFailures.Add(ex.Message);
                                continue;
                            }
                            if(mixedFrame)
                                Notice("EXTERIOR_FRAME_RECOVERED","Информация","Расчётная рамка отделена от границ здания. Обратные служебные участки удалены; замкнутые контуры восстановлены без добавления отрезков.",floor,indicator.ToString());
                            if(RegionIntersection(native,occupied).Area*.09290304<=.001)
                            {
                                try
                                {
                                    var material=FloorMaterialRegion(physicalCandidates,elevation,indicator,native);
                                    if(RegionDifference(native,material).Area*.09290304>.005)
                                        throw new UnconfirmedEnvelopeException("Замкнутая область площадью "+(native.Area*.09290304).ToString("0.###")+" м² не содержит помещений и не подтверждена как отдельная конструкция. Она не включена в здание автоматически.");
                                    Notice("DETACHED_CONSTRUCTION","Информация","Отдельная область без помещений целиком занята конструкцией и не образует площадь здания: "+(native.Area*.09290304).ToString("0.###")+" м².",floor,indicator.ToString());
                                }
                                catch(OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                catch(Exception ex){exteriorFailures.Add(ex.Message);nativeRegions.Add(native);}
                                continue;
                            }
                            if(interior)
                            {
                                try
                                {
                                    if(nativeCurves.All(c=>c is Line))
                                    {
                                        var offsets=new List<double?>();var failures=new List<string>();
                                        foreach(var segment in segments)
                                        {
                                            try{var boundary=ResolveShellBoundary(segment,physicalCandidates,elevation,floor,indicator);offsets.Add(PhysicalBoundaryInset(boundary,segment.GetCurve(),elevation,indicator));}
                                            catch(OperationCanceledException){throw;}
                                            catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                                            catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                            catch(Exception ex){offsets.Add(null);failures.Add(ex.Message);}
                                        }
                                        try{innerBands=RegionUnion(innerBands,native.BoundaryBandsFromSegments(nativeCurves.Select(c=>CurvePoints(c).ToArray()).ToList(),offsets));}
                                        catch(OperationCanceledException){throw;}
                                        catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                                        catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                        catch(Exception ex){throw new UnconfirmedEnvelopeException(ex.Message+" Причины недоступных ограждений: "+string.Join(" | ",failures.Take(2)));}
                                    }
                                    else
                                    {
                                        if(mixedFrame)throw new UnconfirmedEnvelopeException("Наружные криволинейные границы отделены от рамки, но внутренняя сторона объединённой петли требует проверки. Предварительно сохранена наружная граница.");
                                        var offsets=segments.Select(segment=>PhysicalBoundaryInset(ResolveShellBoundary(segment,physicalCandidates,elevation,floor,indicator),segment.GetCurve(),elevation,indicator)).ToList();
                                        innerBands=RegionUnion(innerBands,BoundaryBands(nativeCurves,offsets));
                                    }
                                }
                                catch(OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                catch(Exception ex){interiorFailures.Add(ex.Message);}
                            }
                            else foreach(var segment in segments)
                            {
                                // A cancelled straight seam has no interval on the normalized
                                // boundary and must not be mistaken for a physical shell edge.
                                if(mixedFrame&&segment.GetCurve() is Line&&
                                    !native.HasBoundaryOverlap(CurvePoints(segment.GetCurve()),4/PlanarRegion.Scale))continue;
                                try{ReadNativeShellBoundary(segment,physicalCandidates,floor,indicator);}
                                catch(OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                catch(Exception ex){exteriorFailures.Add(ex.Message);}
                            }
                            nativeRegions.Add(native);
                        }
                        if (nativeRegions.Count == 0&&frameFailures.Count==0&&exteriorFailures.Count==0) throw new InvalidOperationException("Внутри внешнего помещения нет замкнутой оболочки здания.");


                    }
                    finally
                    {
                        if (transaction.GetStatus() == TransactionStatus.Started && transaction.RollBack() != TransactionStatus.RolledBack)
                            throw new InvalidOperationException("Не удалось откатить временное внешнее помещение.");
                        DisposeSpatialCalculators(); localSpatialVolumes.Clear(); parameterCache.Clear(); floorWallFaces.Clear(); floorFaceHeights.Clear();
                    }
                }
                var result=nativeRegions.Aggregate(PlanarRegion.Empty,RegionUnion);
                if(frameFailures.Count>0)
                    throw new ExteriorFrameConflictException(string.Join(" | ",frameFailures.Distinct()),result,
                        "Сохранены только замкнутые части наружной границы. Полный контур этажа не подтверждён; предварительная область не включена в расчёт.");
                if(exteriorFailures.Count>0)
                    throw new UnconfirmedEnvelopeException(string.Join(" | ",exteriorFailures.Distinct()),result,
                        "Контур построен, но происхождение части границ не подтверждено. Предварительная область сохранена для проверки, в итог ТЭП не включена.");
                // Complete all native rings before exposing a failed inward measurement as a draft.
                // Its geometry remains the exterior boundary; no unknown thickness is invented.
                if(interiorFailures.Count>0)
                    throw new UnconfirmedEnvelopeException(string.Join(" | ",interiorFailures.Distinct()),result,
                        "Основа по внешней границе. Для внутреннего контура исправьте границы по внутренним поверхностям наружных стен; внутренний обмер не подтверждён.");
                var measured=interior?RegionDifference(result,innerBands):result;
                // Check BOTH omitted rooms and unclassified space. A closed fragment must
                // not be published as the whole floor after a frame or shell failure.
                try
                {
                    var occupiedRooms=rooms.Aggregate(shaftAreas,(r,item)=>RegionUnion(r,RoomNetRegion(item)));
                    var allMaterial=FloorMaterialRegion(physicalCandidates,elevation,indicator,result);
                    ValidateExteriorCoverage(result,occupiedRooms,allMaterial);
                    // Inward wall bands must not cut away any selected room or shaft.
                    if(interior)ValidateExteriorCoverage(measured,occupiedRooms,allMaterial,true);
                }
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex)
                {
                    throw new UnconfirmedEnvelopeException(ex.Message,measured,(ex as UnconfirmedEnvelopeException)?.CandidateDescription??
                        "Контур построен, но проверка его состава не завершена. Предварительная область сохранена для исправления, в итог ТЭП не включена.");
                }
                result=measured;
                if (result.IsEmpty) throw new InvalidOperationException("Пустая плоская оболочка этажа.");
                Notice("EXTERIOR_ROOM", "Информация", "Граница определена внешним помещением на отметке пола по всем стенам выбранных источников. " +
                    (interior ? "Внутренняя площадь получена 2D-вычитанием полос ограждений и их стыков, без смещения целой петли. " : "") +
                    "Площадь и вычитания вычисляются в 2D без построения вспомогательного тела. Допуск аппроксимации кривых до 0,001 м²; координатная сетка 0,001 мм. Временные объекты отменены транзакцией.", floor, indicator.ToString());
                return result;
            }
        }
    }
}

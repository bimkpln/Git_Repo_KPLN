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
            private readonly Dictionary<string, List<Solid>> roomFloorContours = new Dictionary<string, List<Solid>>();
            private readonly HashSet<string> inferredRoomFloorContours = new HashSet<string>();
            private readonly Dictionary<string, string> roomFloorFailures = new Dictionary<string, string>();
            private readonly HashSet<double> invalidRoomFloors = new HashSet<double>();

            private void MarkInvalidRoomFloor(Source source, Element element)
            {
                var room = element as Room;
                if (room?.Level != null)
                    invalidRoomFloors.Add(Math.Round(source.Transform.OfPoint(new XYZ(0, 0, room.Level.ProjectElevation)).Z, 6));
            }

            private static bool RoomEnvelopeTotal(Indicator indicator)
            {
                return OneOf(indicator.ToString(), "Gns", "GnsResidential", "GnsNonresidential",
                    "Gross", "GrossAbove", "GrossBelow", "Np", "NpResidential", "NpNonresidential");
            }

            private static bool RoomEnvelopePart(Indicator indicator)
            {
                return OneOf(indicator.ToString(), "GnsLivingPart", "GnsNonlivingPart", "NnpEmbedded", "NnpSeparate");
            }

            public void PrepareRoomWorkflow()
            {
                // The old rule table is retained in the backup, never applied invisibly behind the cards.
                if (Config.RoomWorkflowVersion != 4)
                {
                    Config.PreviousWorkflowSettings = Serialize(Config);
                    Config.RoomWorkflowVersion = 4;
                }
                Config.Departments = Config.Departments ?? new List<DepartmentAssignment>();
                Config.Departments.RemoveAll(d=>Eq(d.Value,"Квартира"));
                Config.Rules.Clear();
                Config.CreateViews = false;
                foreach(var map in Config.Parameters.Where(m=>Settings.IsFixedParameter(m.Key)))
                    map.Name=Settings.FixedParameterName(map.Key);
                Config.ZeroMode = "level";
                Config.ZeroValue = ""; Config.ZeroParameter = "";
                Config.Geometry = "rooms";
                Config.ApartmentMode = "id";
                Config.ParkingMode = "id";
                Config.AreaScheme = null;
                Config.Grouping = "parameter";
                Config.Buildings.Clear();
                Config.Corrections.Clear();
                Config.Contours.Clear();
                foreach (var source in Sources) source.Mode = "include";
                foreach (var metric in Config.Metrics)
                {
                    var indicator = (Indicator)Enum.Parse(typeof(Indicator), metric.Key);
                    metric.Geometry = "rooms";
                    metric.AreaScheme = null;
                    metric.ContourMode = RoomEnvelopeTotal(indicator) ? "walls" : "current";
                    metric.WallSelectionMode = "exterior";
                    metric.WallBoundary = "auto";
                    metric.WallCutHeight = "0";
                    metric.ContourExclusions = "normative";
                    metric.SelectedWalls.Clear();
                }
            }

            private string RoomFloorKey(double elevation, Indicator indicator)
            {
                return elevation.ToString("R", CultureInfo.InvariantCulture) + "/" +
                    ((int)indicator <= 4 ? Methodology.GnsWallBoundary : "interior");
            }

            private List<Solid> RoomFloorContours(List<Record> records, Record floor, Indicator indicator)
            {
                double cut = RoomFloorElevation(floor);
                var peers = records.Where(r => r.Element is Room && r.Level != null && r.Level.Include && Math.Abs(r.Z - floor.Z) < 1e-6).ToList();
                if (peers.Any(r => Math.Abs(RoomFloorElevation(r) - cut) > FloorGeometryTolerance))
                    throw new InvalidOperationException("Помещения одного расчётного этажа имеют разные отметки пола. Общий контур нельзя измерять на одной произвольной высоте; требуется разделение участков перепада.");
                string key = RoomFloorKey(cut, indicator);
                List<Solid> contours;
                string failure;
                if (roomFloorFailures.TryGetValue(key, out failure)) throw new InvalidOperationException(failure);
                if (invalidRoomFloors.Contains(Math.Round(floor.Z, 6)))
                    throw new InvalidOperationException("На этаже есть размещённые помещения с ошибками площади или классификации. Полный контур этажа не рассчитан; см. ошибки помещений.");
                if (roomFloorContours.TryGetValue(key, out contours))
                {
                    if(inferredRoomFloorContours.Contains(key))Notice("EXTERIOR_BY_GEOMETRY","Предупреждение","Повторно использована оболочка, подтверждённая внешним помещением без признака «Наружная» у стен. Отметка пола не изменялась.",floor,indicator.ToString());
                    return contours;
                }
                try
                {
                    // Model walls, including linked walls, are geometry sources. Room parameters own classification.
                    var allWalls = Sources.Where(s => s.Loaded && s.LoadError == null && s.Mode == "include")
                        .SelectMany(s => s.Elements.OfType<Wall>().SelectMany(WallMembers).Distinct()
                            .Where(w => PhaseAccepted(s, w))
                            .Select(w => new Record { Source = s, Element = w, Level = floor.Level, Role = "structure" }))
                        .Where(r => CrossesFloor(r.Element, r.Source.Transform, cut)).ToList();
                    var walls=allWalls.Where(r=>((Wall)r.Element).WallType.Function==WallFunction.Exterior).ToList();
                    bool inferred=walls.Count==0;
                    if(inferred)
                    {
                        if(allWalls.Count==0)throw new InvalidOperationException("На отметке пола "+(cut*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м нет стен выбранной стадии. Высота сечения не изменялась.");
                        if((int)indicator>4)throw new InvalidOperationException("На отметке пола есть стены ("+allWalls.Count+"), но ни одна не имеет функцию «Наружная». Внутренняя граница общей площади требует подтверждения стороны оболочки; произвольная сторона стены не используется.");
                        // Remove only walls proven to be shared by two placed rooms. Every remaining
                        // candidate must be seen by the exterior room, including courtyard candidates.
                        var counts=new Dictionary<string,int>();
                        foreach(var peer in peers)
                        {
                            var used=new HashSet<string>();
                            var rings=((Room)peer.Element).GetBoundarySegments(new SpatialElementBoundaryOptions{SpatialElementBoundaryLocation=SpatialElementBoundaryLocation.Finish});
                            if(rings==null)continue;
                            foreach(var segment in rings.SelectMany(r=>r))
                            {
                                var wall=peer.Element.Document.GetElement(segment.ElementId) as Wall;
                                if(wall==null)continue; // Linked/unresolved boundaries remain candidates, never discarded.
                                foreach(var member in WallMembers(wall))used.Add(peer.Source.Key+"/"+member.UniqueId);
                            }
                            foreach(var wallKey in used){int count;counts.TryGetValue(wallKey,out count);counts[wallKey]=count+1;}
                        }
                        walls=allWalls.Where(r=>!counts.ContainsKey(r.Key)||counts[r.Key]<2).ToList();
                        if(walls.Count==0)throw new InvalidOperationException("После исключения общих стен помещений не осталось кандидатов внешней оболочки.");
                        Notice("EXTERIOR_BY_GEOMETRY","Предупреждение","Стены с функцией «Наружная» не найдены. ГНС проверяется по границе внешнего помещения на той же отметке пола. Разрывы, невидимые кандидаты и возможные дворы сохраняют отказ расчёта.",floor,indicator.ToString());
                    }
                    Progress("2D-контур наружных стен: " + floor.Level.Name + "; стен: " + walls.Count);
                    var metric = new Metric { WallBoundary = "auto" };
                    Solid envelope;
                    if ((int)indicator <= 4)
                    {
                        try { envelope = ExteriorRoomContour(walls, floor, indicator, cut, inferred); }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            if(inferred)throw new InvalidOperationException("Геометрически определяемая оболочка не подтверждена: "+ex.Message,ex);
                            Notice("EXTERIOR_ROOM_FALLBACK", "Информация", "Внешнее помещение: " + ex.Message + " Используется резервное сечение стен на той же отметке.", floor, indicator.ToString());
                            envelope = WallContour(walls, metric, indicator, cut);
                        }
                    }
                    else envelope = WallContour(walls, metric, indicator, cut);
                    contours = PlanPieces(envelope).ToList();
                    if (contours.Count == 0) throw new InvalidOperationException("Не получен замкнутый контур этажа.");
                    roomFloorContours.Add(key, contours);
                    if(inferred)inferredRoomFloorContours.Add(key);
                    return contours;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    failure = "Этаж «" + floor.Level.Name + "»: " + ex.Message;
                    roomFloorFailures[key] = failure;
                    throw new InvalidOperationException(failure, ex);
                }
            }

            private Solid RoomNetPlan(Record room)
            {
                return RoomBoundaryPlan(room, "net");
            }

            private Solid RoomBoundaryPlan(Record room, string boundary, string metric = "")
            {
                string key = RoomPlanCacheKey(room, boundary, metric);
                Solid shape;
                if (!shapes.TryGetValue(key, out shape)) shapes[key] = shape = SpatialPlan(room, boundary, metric);
                return shape;
            }

            private Solid ClipRoomPart(Solid shape, Record room, List<Record> records, Indicator indicator)
            {
                Solid result = null;
                foreach (var contour in RoomFloorContours(records, room, indicator))
                    result = Union(result, Intersect(shape, contour));
                if (result == null || result.Volume < 1e-9)
                    throw new InvalidOperationException("Помещение не попало внутрь контура наружных стен своего этажа.");
                return result;
            }

            private void CalculateRoomEnvelopeAreas(List<Record> records, Indicator indicator)
            {
                var rooms = records.Where(r => r.Element is Room && r.Level != null && r.Level.Include).ToList();
                if (rooms.Count == 0) throw new InvalidOperationException("Нет размещённых помещений на включённых этажах. Для расчёта требуются Rooms.");
                foreach (var level in Config.Levels.Where(l => l.Include && !OneOf(l.Kind, "exclude", "attic"))
                    .GroupBy(l => Math.Round(l.Elevation, 6)).Select(g => g.First()))
                {
                    if (rooms.Any(r => Math.Abs(r.Z - level.Elevation) < 1e-6)) continue;
                    bool aboveOnly = (int)indicator <= 4 || OneOf(indicator.ToString(), "GrossAbove", "Np", "NpResidential", "NpNonresidential");
                    if (aboveOnly && (level.Above == "below" || level.Above == "auto" && level.Kind == "normal" && level.Elevation < ground)) continue;
                    if (indicator == Indicator.GrossBelow && (level.Above == "above" || level.Above == "auto" && level.Kind == "normal" && level.Elevation >= ground)) continue;
                    current.Issue("ROOM_LEVEL_EMPTY", "Ошибка", "На включённом этаже «" + level.Name + "» нет помещений. Нельзя проверить состав площади. Проверьте состав расчётных этажей.", indicator.ToString(), level.Source);
                }
                foreach (var floor in rooms.GroupBy(r => Math.Round(r.Z, 6)).OrderBy(g => g.Key))
                {
                    var first = floor.First();
                    var rows = new List<Detail>();
                    try
                    {
                        if (!floor.Any(r => ContourBaseIncluded(r, indicator))) continue;
                        if(invalidRoomFloors.Contains(floor.Key))
                        {
                            current.Issue("ROOM_FLOOR_SKIPPED","Ошибка","Этаж «"+string.Join(" / ",floor.Select(r=>r.Level.Name).Distinct())+"» ("+(floor.Key*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" м) пропущен: есть размещённые помещения с ошибками. Полную площадь контура подтвердить нельзя. Остальные этажи рассчитываются.",indicator.ToString(),first.Source.Name,
                                action:"См. ошибки помещений с ID. Пропущенный этаж не включён в сумму и не считается этажом с нулевой площадью.");
                            continue;
                        }
                        int firstIssue = current.Issues.Count;
                        var contours = RoomFloorContours(records, first, indicator);
                        var assigned = new HashSet<string>();
                        foreach (var contour in contours)
                        {
                            var occupants = floor.Where(r => (Intersect(RoomNetPlan(r), contour)?.Volume ?? 0) > 1e-9).ToList();
                            if (occupants.Count == 0)
                                throw new InvalidOperationException("В замкнутом контуре этажа нет помещений. Нельзя определить назначение площади и проверить исключения.");
                            foreach (var room in occupants)
                                if (!assigned.Add(room.Key)) throw new InvalidOperationException("Помещение ID " + IDHelper.ElIdValue(room.Element.Id) + " пересекает несколько отдельных контуров зданий.");
                            if (occupants.Select(r => r.Building + "|" + r.BuildingClass + "|" + r.Profile).Distinct().Count() != 1)
                                throw new InvalidOperationException("В одном замкнутом контуре разные корпуса или профили здания. Проверьте параметры помещений.");
                            if (occupants.Select(r => r.Level.Kind + "|" + r.Level.Above + "|" + r.Level.TopSlab).Distinct().Count() != 1)
                                throw new InvalidOperationException("Помещения одного контура имеют противоречащие настройки наземности / вида этажа.");
                            var owner = occupants.First().Copy();
                            owner.Section = occupants.Select(r => r.Section).Distinct().Count() == 1 ? owner.Section : "";
                            owner.Override = null;
                            if (!ContourBaseIncluded(owner, indicator)) continue;
                            Solid exclusions = null, occupied = null;
                            foreach (var room in occupants)
                            {
                                Progress("2D-помещения: " + owner.Level.Name + "; ID " + IDHelper.ElIdValue(room.Element.Id));
                                var net = RoomNetPlan(room);
                                if ((Intersect(occupied, net)?.Volume ?? 0) * .09290304 > .005)
                                    throw new InvalidOperationException("Пересекающиеся помещения на этаже, ID " + IDHelper.ElIdValue(room.Element.Id) + ". Назначение пересечения неоднозначно.");
                                occupied = Union(occupied, net);
                                if (room.Role == "unknown") throw new InvalidOperationException("Не определено назначение помещения ID " + IDHelper.ElIdValue(room.Element.Id) + ". Задайте параметр или правило классификации.");
                                bool include = Eligible(room, indicator);
                                if (include && (Subtract(net, contour)?.Volume ?? 0) * .09290304 > .005)
                                    throw new InvalidOperationException("Граница учитываемого помещения ID " + IDHelper.ElIdValue(room.Element.Id) + " выходит за контур наружных стен. Проверьте геометрию этажа.");
                                // Ordinary included rooms require no exclusion polygon. Read the
                                // envelope once; do not reconstruct every room's exterior wall share.
                                bool vertical = OneOf(room.Role, "multilight", "stair-gap", "opening", "shaft", "engineering-shaft");
                                bool heightCut = (int)indicator > 4 && (OneOf(room.Role, "under-stair", "niche", "arch") || Number(room.Element, "slope").HasValue);
                                if (include && !vertical && !heightCut) continue;
                                var full = RoomBoundaryPlan(room, (int)indicator <= 4 ? "outer" : "gross", indicator.ToString());
                                var measured = include ? Plan(room, indicator) : null;
                                if (include && VerticalExclusion(room, indicator, records, measured ?? full)) include = false;
                                var mask = include ? (ReferenceEquals(full, measured) ? null : Subtract(full, measured)) : full;
                                mask = Intersect(contour, mask);
                                exclusions = Union(exclusions, mask);
                                if (mask != null && mask.Volume > 1e-9)
                                    rows.Add(Row(room, indicator, mask.Volume * .09290304, 1, true,
                                        "2D-исключение по назначению, высоте или правилу вертикального пространства", mask));
                            }
                            // Openings without Rooms and parametrically designated exceptions retain their checks.
                            foreach (var element in records.Where(r => !(r.Element is SpatialElement) && r.Level != null && Math.Abs(r.Z - owner.Z) < 1e-6))
                            {
                                if (!ContourMaskRequired(element, Config.Metrics.Single(m => m.Key == indicator.ToString()), indicator, records)) continue;
                                var mask = Intersect(contour, ContourMaskPlan(element, indicator));
                                exclusions = Union(exclusions, mask);
                                if (mask != null && mask.Volume > 1e-9)
                                    rows.Add(Row(element, indicator, mask.Volume * .09290304, 1, true, "2D-исключение элемента", mask));
                            }
                            var result = Subtract(contour, exclusions);
                            if (result != null && result.Volume > 1e-9)
                            {
                                var row = Row(owner, indicator, result.Volume * .09290304, 1, false,
                                    "Контур наружных стен на отметке пола " + (RoomFloorElevation(owner) * .3048).ToString("0.######", CultureInfo.InvariantCulture) + " м минус объединение 2D-исключений; внутренние стены учтены", result);
                                row.Purpose = "envelope";
                                rows.Add(row);
                            }
                        }
                        if (floor.Any(r => !assigned.Contains(r.Key) && Eligible(r, indicator)))
                            throw new InvalidOperationException("Есть помещения вне замкнутых контуров наружных стен. Проверьте наружные стены и привязку этажа.");
                        if (current.Issues.Skip(firstIssue).Any(i => i.Severity == "Ошибка"))
                            throw new InvalidOperationException("Проверки исключений этажа завершились ошибкой. Площадь этажа не опубликована; см. связанные сообщения.");
                        // An invalid floor must never publish the successfully processed subset as its total.
                        current.Details.AddRange(rows);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("ROOM_FLOOR_CONTOUR", "Ошибка", first.Level.Name + ": " + ex.Message, first, indicator.ToString()); }
                }
            }
        }
    }
}

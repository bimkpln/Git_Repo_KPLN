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
            private readonly Dictionary<string,List<Issue>> roomFloorWarnings=new Dictionary<string,List<Issue>>();
            private readonly Dictionary<string, List<PlanarRegion>> roomFloorRegions = new Dictionary<string, List<PlanarRegion>>();
            private List<Record> contourRoomRecords = new List<Record>();
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
                // Obsolete rules are not retained as recursively embedded settings snapshots.
                Config.PreviousWorkflowSettings=null;
                Config.RoomWorkflowVersion=4;
                Config.Departments = Config.Departments ?? new List<DepartmentAssignment>();
                Config.Rules.Clear();
                Config.CreateViews = false;Config.UseReviewRegions=true;
                Config.ClassifyFamilies=false;Config.FamilyAreaFromParameter=false;Config.RoomClassificationParameter="@Department";
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
                string key = RoomFloorKey(RoomFloorElevation(floor), indicator);
                List<Solid> result;
                if (!roomFloorContours.TryGetValue(key, out result))
                    roomFloorContours[key] = result = RoomFloorRegions(records, floor, indicator).Select(r => BuildPlanarSolid(r, 0, 1, true)).ToList();
                return result;
            }

            private List<PlanarRegion> RoomFloorRegions(List<Record> records, Record floor, Indicator indicator)
            {
                double cut = RoomFloorElevation(floor);
                var peers = (contourRoomRecords.Count > 0 ? contourRoomRecords : records)
                    .Where(r => IsAreaInput(r.Element) && r.Level != null && r.Level.Include && Math.Abs(r.Z - floor.Z) < 1e-6).ToList();
                if (peers.Any(r => Math.Abs(RoomFloorElevation(r) - cut) > FloorGeometryTolerance))
                    throw new InvalidOperationException("Помещения одного этажа имеют разные отметки пола. Требуется разделение участков перепада; произвольная высота не используется.");
                string key = RoomFloorKey(cut, indicator);
                List<PlanarRegion> result; string failure;
                if (roomFloorRegions.TryGetValue(key, out result))
                {
                    List<Issue> warnings;
                    if(roomFloorWarnings.TryGetValue(key,out warnings))foreach(var warning in warnings)
                        Notice(warning.Code,warning.Severity,warning.Message,floor,indicator.ToString());
                    return result;
                }
                if (roomFloorFailures.TryGetValue(key, out failure)) throw new InvalidOperationException(failure);
                try
                {
                    var envelope = ReviewRegion("floor/"+key, floor, "Контур этажа - "+((int)indicator<=4?Methodology.GnsWallBoundary:"interior"),cut,()=>
                    {
                    var walls = Sources.Where(s => s.Loaded && s.LoadError == null && s.Mode == "include")
                        .SelectMany(s => s.Elements.OfType<Wall>().SelectMany(WallMembers).Distinct()
                            .Where(w => PhaseAccepted(s, w))
                            .Select(w => new Record { Source = s, Element = w, Level = floor.Level, Role = "structure" }))
                        .Where(r => CrossesFloor(r.Element, r.Source.Transform, cut)).ToList();
                    if (walls.Count == 0) throw new InvalidOperationException("На отметке пола нет стен выбранных источников и стадии.");
                    Progress("Плоская оболочка: " + floor.Level.Name + "; стен: " + walls.Count);
                    if(Config.Phase==Settings.AllPhases)
                    {
                        // A temporary Revit Room belongs to one phase. For an unfiltered set use actual wall sections.
                        var metric=Config.Metrics.First(m=>m.Key==indicator.ToString());
                        var exteriorWalls=walls.Where(r=>ContourWallSelected(r,metric)).ToList();
                        if(exteriorWalls.Count==0)throw new InvalidOperationException("Не найдены наружные стены для контура всех стадий.");
                        var allStages=WallContourRegion(exteriorWalls,metric,indicator,cut);
                        var occupied=peers.Aggregate(AutomaticShaftFloorRegion(cut),(region,room)=>RegionUnion(region,RoomNetRegion(room)));
                        if(RegionDifference(occupied,allStages).Area*.09290304>.005)
                            throw new UnconfirmedEnvelopeException("Контур всех стадий не охватывает выбранные помещения. Проверьте совмещённые состояния модели или выберите конкретную стадию.",allStages,
                                "Контур построен по выбранным стенам всех стадий, но не охватывает помещения. Область сохранена для проверки и исправления, в итог ТЭП не включена.");
                        var material=FloorMaterialRegion(PhysicalShellCandidates(walls,floor,cut),cut,indicator,allStages);
                        if(RegionDifference(allStages,RegionUnion(occupied,material)).Area*.09290304>.005)
                            throw new UnconfirmedEnvelopeException("В контуре всех стадий есть область без помещений и подтверждённых конструкций. Её назначение не определено.",allStages,
                                "Контур построен, но содержит неподтверждённую площадь. Область сохранена для проверки и исправления, в итог ТЭП не включена.");
                        return allStages;
                    }
                    try{return ExteriorRoomRegion(walls, floor, indicator, cut, peers);}
                    catch(ExteriorFrameConflictException ex)
                    {
                        // ExteriorRoomRegion has already rolled back every temporary object.
                        return RecoverExteriorFrameRegion(walls,floor,indicator,cut,peers,ex);
                    }
                    });
                    int issueStart=current.Issues.Count;
                    result = envelope.Components().Select(paths => new PlanarRegion(paths)).ToList();
                    if (result.Count == 0) throw new InvalidOperationException("Не получена замкнутая оболочка этажа.");
                    roomFloorWarnings[key]=current.Issues.Skip(issueStart).Where(i=>i.Severity=="Предупреждение").ToList();
                    roomFloorRegions[key] = result;
                    return result;
                }
                catch (OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
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

            private sealed class FloorCell
            {
                internal PlanarRegion Region;
                internal List<Record> Owners;
            }
            private readonly Dictionary<string,PlanarRegion> roomCentreRegions=new Dictionary<string,PlanarRegion>();
            private PlanarRegion RoomCentreRegion(Record room)
            {
                PlanarRegion edited;if(TryEditedObjectOverlay(room,out edited))return edited;
                if(IsClassifiedFamily(room.Element))return FamilyAreaRegion(room);
                string key=RoomPlanCacheKey(room,"centre","");PlanarRegion region;
                if(!roomCentreRegions.TryGetValue(key,out region))
                {
                    region=ReviewRegion("centre/"+key,room,"Помещение - по осям для СПП",RoomFloorElevation(room),()=>SectionRegion(SpatialVolume(room,true),RoomFloorElevation(room)));
                    if(region.IsEmpty)throw new InvalidOperationException("Нет сечения Room по осям ограждений на отметке пола, ID "+IDHelper.ElIdValue(room.Element.Id));
                    roomCentreRegions[key]=region;
                }
                return region;
            }
            private List<FloorCell> PartitionRoomFloor(PlanarRegion envelope,List<Record> rooms)
            {
                var regions=rooms.Select(room=>{
                    Progress("2D-разделение этажа: ID "+IDHelper.ElIdValue(room.Element.Id));return RoomCentreRegion(room);
                }).ToList();
                return PlanarRegion.PartitionByRooms(envelope,regions,planarCheckpoint)
                    .Select(cell=>new FloorCell{Region=cell.Region,Owners=cell.Owners.Select(index=>rooms[index]).ToList()}).ToList();
            }
            private static Indicator EnvelopeBaseIndicator(Indicator metric)
            {
                if(metric==Indicator.GnsLivingPart||metric==Indicator.GnsNonlivingPart)return Indicator.GnsResidential;
                if(metric==Indicator.NnpEmbedded||metric==Indicator.NnpSeparate)return Indicator.GrossAbove;
                return metric;
            }
            private PlanarRegion FloorCellExclusion(FloorCell cell,Indicator metric,List<Record> records,List<FloorCell> cells)
            {
                var included=new List<bool>();
                foreach(var room in cell.Owners)
                {
                    bool value=Eligible(room,metric);
                    double area=cells.Where(c=>c.Owners.Count==1&&c.Owners[0].Key==room.Key).Sum(c=>c.Region.Area);
                    if(value&&room.Role=="opening"&&(int)metric<=4)
                    {
                        double maximum=area+cells.Where(c=>c.Owners.Count>1&&c.Owners.Any(r=>r.Key==room.Key)).Sum(c=>c.Region.Area);
                        if(Methodology.CountsGnsOpeningAreaOnOneFloor(area*.09290304)!=Methodology.CountsGnsOpeningAreaOnOneFloor(maximum*.09290304))
                            throw new InvalidOperationException("Общая конструкция влияет на порог площади проёма 36 м², ID "+IDHelper.ElIdValue(room.Element.Id)+". Нормативное исключение без разделения не подтверждено.");
                    }
                    if(value&&VerticalExclusionArea(room,metric,records,area))value=false;
                    included.Add(value);
                }
                if(included.Distinct().Count()>1)
                    throw new InvalidOperationException("Участок конструкций "+(cell.Region.Area*.09290304).ToString("0.######",CultureInfo.InvariantCulture)+" м² граничит с включаемыми и исключаемыми помещениями. Для этого показателя разделение не определено: ID "+string.Join(", ",cell.Owners.Select(r=>IDHelper.ElIdValue(r.Element.Id)))+". Площадь не распределена пропорционально или по ближайшему помещению.");
                if(!PlanarRegion.CommonCellInclusion(included))return cell.Region;
                bool heightCut=(int)metric>4&&cell.Owners.Any(r=>OneOf(r.Role,"under-stair","niche","arch")||NeedsCeilingCheck(r,metric)&&HasSlopedCeiling(r));
                if(!heightCut)return PlanarRegion.Empty;
                if(cell.Owners.Count!=1)throw new InvalidOperationException("Для общего участка конструкций требуется отдельное подтверждение высотного исключения.");
                return RegionDifference(cell.Region,ApplyRoomHeightRules(cell.Owners[0],metric,cell.Region,true));
            }

            private void CalculateRoomEnvelopeAreas(List<Record> records, Indicator indicator)
            {
                var rooms = records.Where(r => IsAreaInput(r.Element) && r.Level != null && r.Level.Include && r.Source.Mode=="include").ToList();
                var baseIndicator=EnvelopeBaseIndicator(indicator);
                if (rooms.Count == 0) throw new InvalidOperationException("Нет размещённых помещений на включённых этажах. Нужны размещённые Rooms или классифицированные экземпляры с геометрией.");
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
                        if (!floor.Any(r => ContourBaseIncluded(r, baseIndicator))) continue;
                        int firstIssue = current.Issues.Count;
                        var contours = RoomFloorRegions(records, first, indicator);
                        if (invalidRoomFloors.Contains(floor.Key))
                        {
                            double area = contours.Sum(c => c.Area) * .09290304;
                            Notice("FLOOR_ENVELOPE_ONLY", "Информация", "Этаж «" + first.Level.Name + "»: площадь подтверждённой оболочки до исключений " + area.ToString("0.###", CultureInfo.InvariantCulture) + " м². Это справочная геометрия, в итог ТЭП не включена: состав исключений не подтверждён.", first, indicator.ToString());
                            foreach (var contour in contours)
                            {
                                var reference = PlanarRow(first, indicator, contour, true, "Справочная оболочка до исключений; не является итогом ТЭП");
                                reference.Purpose = "envelope-reference";
                                current.Details.Add(reference);
                            }
                            Notice("ROOM_FLOOR_SKIPPED", "Ошибка", "Этаж «" + first.Level.Name + "»: оболочка построена независимо от ошибки помещения, но итог после исключений не подтверждён. Другие этажи рассчитываются.", first, indicator.ToString());
                            continue;
                        }
                        var assigned = new HashSet<string>();
                        foreach (var contour in contours)
                        {
                            var occupants = floor.Where(r => RegionIntersection(RoomNetRegion(r), contour).Area > 1e-9).ToList();
                            if (occupants.Count == 0)
                                throw new InvalidOperationException("В замкнутом контуре этажа нет помещений. Нельзя определить назначение площади и проверить исключения.");
                            foreach (var room in occupants)
                                if (!assigned.Add(room.Key)) throw new InvalidOperationException("Помещение ID " + IDHelper.ElIdValue(room.Element.Id) + " пересекает несколько отдельных контуров зданий.");
                            if (occupants.Select(r => r.Building + "|" + r.BuildingClass + "|" + r.Profile).Distinct().Count() != 1)
                                throw new InvalidOperationException("В одном замкнутом контуре разные корпуса или профили здания. Проверьте параметры помещений.");
                            if (occupants.Select(r => r.Level.Kind + "|" + r.Level.Above + "|" + TopSlabRuleKey(r.Level)).Distinct().Count() != 1)
                                throw new InvalidOperationException("Помещения одного контура имеют противоречащие настройки наземности / вида этажа.");
                            var owner = occupants.First().Copy();
                            owner.Section = occupants.Select(r => r.Section).Distinct().Count() == 1 ? owner.Section : "";
                            owner.Override = null;
                            if (!ContourBaseIncluded(owner, baseIndicator)) continue;
                            var exclusions=PlanarRegion.Empty;var occupied=PlanarRegion.Empty;
                            foreach(var room in occupants)
                            {
                                var net=RoomNetRegion(room);
                                if(RegionIntersection(occupied,net).Area*.09290304>.005)throw new InvalidOperationException("Пересекающиеся помещения, ID "+IDHelper.ElIdValue(room.Element.Id));
                                occupied=RegionUnion(occupied,net);
                                if(room.Role=="unknown")throw new InvalidOperationException("Не распределено назначение помещения ID "+IDHelper.ElIdValue(room.Element.Id));
                                if(RegionDifference(net,contour).Area*.09290304>.005)throw new InvalidOperationException("Помещение выходит за нормативный контур, ID "+IDHelper.ElIdValue(room.Element.Id));
                            }
                            var cells=PartitionRoomFloor(contour,occupants);
                            foreach(var cell in cells)
                            {
                                var mask=FloorCellExclusion(cell,indicator,records,cells);
                                if(mask.IsEmpty)continue;
                                exclusions=RegionUnion(exclusions,mask);
                                rows.Add(PlanarRow(cell.Owners[0],indicator,mask,true,"2D-исключение; сечение Room по осям ограждений и подтверждённые смежные конструкции; ID помещений: "+string.Join(", ",cell.Owners.Select(r=>IDHelper.ElIdValue(r.Element.Id)))));
                            }
                            // Openings without Rooms and parametrically designated exceptions retain their checks.
                            foreach (var element in records.Where(r => !IsAreaInput(r.Element) && r.Level != null && Math.Abs(r.Z - owner.Z) < 1e-6))
                            {
                                if (!ContourMaskRequired(element, Config.Metrics.Single(m => m.Key == indicator.ToString()), indicator, records)) continue;
                                var mask = RegionIntersection(contour, RegionOfPlan(ContourMaskPlan(element, indicator)));
                                var uniqueMask = RegionDifference(mask, exclusions);
                                exclusions = RegionUnion(exclusions, mask);
                                if (!uniqueMask.IsEmpty)
                                    rows.Add(PlanarRow(element, indicator, uniqueMask, true, "2D-исключение элемента; пересечения вычтены один раз"));
                            }
                            var result = RegionDifference(contour, exclusions);
                            double rawArea=contour.Area*.09290304,excludedArea=exclusions.Area*.09290304,netArea=result.Area*.09290304;
                            if(Math.Abs(rawArea-excludedArea-netArea)>.001)throw new InvalidOperationException("Не сошёлся баланс площади оболочки, исключений и остатка с допуском 0,001 м².");
                            Notice("FLOOR_AREA_BALANCE","Информация","Этаж «"+owner.Level.Name+"»: оболочка "+rawArea.ToString("0.###",CultureInfo.InvariantCulture)+" м² - исключения "+excludedArea.ToString("0.###",CultureInfo.InvariantCulture)+" м² = "+netArea.ToString("0.###",CultureInfo.InvariantCulture)+" м². Пересечения исключений учтены один раз.",owner,indicator.ToString());
                            {

                                var row = PlanarRow(owner, indicator, result, false,
                                    "Контур наружных стен на отметке пола " + (RoomFloorElevation(owner) * .3048).ToString("0.######", CultureInfo.InvariantCulture) + " м минус объединение 2D-исключений; внутренние стены учтены");
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
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                    catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                    catch (Exception ex) { Notice("ROOM_FLOOR_CONTOUR", "Ошибка", first.Level.Name + ": " + ex.Message, first, indicator.ToString()); }
                }
            }
        }
    }
}

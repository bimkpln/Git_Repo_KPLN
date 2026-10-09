using Autodesk.Revit.DB;
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
            private sealed class ShaftVoidEvidence
            {
                internal PlanarRegion Region;
                internal string Description;
            }

            private sealed class ShaftPhysicalEvidence
            {
                internal readonly List<ShaftVoidEvidence> Voids = new List<ShaftVoidEvidence>();
                internal readonly List<string> Errors = new List<string>();
                // Includes physical vertical barriers: walls, columns, curtain components
                // and room-bounding families. Equipment bounds never become material.
                internal PlanarRegion WallMaterial = PlanarRegion.Empty;
                internal bool WallsComplete = true;
            }

            private sealed class ShaftBoundsRead
            {
                internal double[] Bounds;
                internal string Error;
            }

            private sealed class ShaftSectionRead
            {
                internal PlanarRegion Region;
                internal string Error;
            }

            // These caches belong to one collection, not to the document lifetime.
            private readonly Dictionary<string, ShaftBoundsRead> shaftEvidenceBounds = new Dictionary<string, ShaftBoundsRead>();
            private readonly Dictionary<string, ShaftSectionRead> shaftEvidenceSections = new Dictionary<string, ShaftSectionRead>();
            private readonly Dictionary<string, ShaftSectionRead> shaftEvidenceOpenings = new Dictionary<string, ShaftSectionRead>();

            private void ClearShaftEvidence()
            {
                shaftEvidenceBounds.Clear();
                shaftEvidenceSections.Clear();
                shaftEvidenceOpenings.Clear();
            }

            private static string ShaftEvidenceKey(Source source, Element element)
            { return source.Key + "/" + element.UniqueId; }

            private static string ShaftEvidenceDescription(Source source, Element element)
            { return source.Name + "; ID " + IDHelper.ElIdValue(element.Id).ToString(CultureInfo.InvariantCulture); }

            private static string ShaftEvidenceElevation(double elevation)
            { return (elevation * .3048).ToString("+0.000;-0.000;0.000", CultureInfo.InvariantCulture) + " м"; }

            private ShaftBoundsRead ReadShaftEvidenceBounds(Source source, Element element)
            {
                string key = ShaftEvidenceKey(source, element);
                ShaftBoundsRead result;
                if (shaftEvidenceBounds.TryGetValue(key, out result)) return result;
                result = new ShaftBoundsRead();
                try { result.Bounds = ElementBounds(element, source.Transform); }
                catch (OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                catch (Exception ex) { result.Error = ex.Message; }
                shaftEvidenceBounds[key] = result;
                return result;
            }

            private static bool ShaftEvidenceBoundsOverlap(double[] bounds, double[] search)
            {
                // Broad phase only. No part of these rectangles becomes a measured contour.
                const double tolerance = .001 / .3048;
                return bounds[3] >= search[0] - tolerance && bounds[0] <= search[2] + tolerance &&
                    bounds[4] >= search[1] - tolerance && bounds[1] <= search[3] + tolerance;
            }

            private ShaftSectionRead ReadShaftMaterialSection(Source source, Element element, double elevation)
            {
                string key = ShaftEvidenceKey(source, element) + "/" + elevation.ToString("R", CultureInfo.InvariantCulture);
                ShaftSectionRead result;
                if (shaftEvidenceSections.TryGetValue(key, out result)) return result;
                result = new ShaftSectionRead();
                try
                {
                    Progress("Проверка границ шахт: сечение " + ShaftEvidenceDescription(source, element));
                    var solids = Solids(element);
                    if (solids.Count == 0) throw new InvalidOperationException("Нет доступных физических тел конструкции.");
                    var region = PlanarRegion.Empty;
                    foreach (var solid in solids)
                        region = RegionUnion(region, SectionRegion(SolidUtils.CreateTransformed(solid, source.Transform), elevation));
                    result.Region = region;
                }
                catch (OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                catch (Exception ex) { result.Error = ex.Message; }
                shaftEvidenceSections[key] = result;
                return result;
            }

            private static double? ShaftOpeningLength(Opening opening, BuiltInParameter name)
            {
                var parameter = opening.get_Parameter(name);
                if (parameter == null || !parameter.HasValue || parameter.StorageType != StorageType.Double) return null;
                double value = parameter.AsDouble();
                return double.IsNaN(value) || double.IsInfinity(value) ? (double?)null : value;
            }

            private static Level ShaftOpeningLevel(Opening opening, BuiltInParameter name)
            {
                var parameter = opening.get_Parameter(name);
                return parameter != null && parameter.HasValue && parameter.StorageType == StorageType.ElementId
                    ? opening.Document.GetElement(parameter.AsElementId()) as Level : null;
            }

            private Tuple<double, double> ShaftOpeningRange(Source source, Opening opening, double[] bounds)
            {
                // Shaft Opening uses the wall constraint parameters in supported Revit versions.
                // If not exposed, the opening's own extent can still constrain its Z range.
                BuiltInCategory shaftCategory;
                bool shaft = Enum.TryParse("OST_ShaftOpening", out shaftCategory) && opening.Category != null &&
                    IDHelper.ElIdValue(opening.Category.Id) == (long)shaftCategory;
                if (shaft)
                {
                    var bottomLevel = ShaftOpeningLevel(opening, BuiltInParameter.WALL_BASE_CONSTRAINT);
                    var topLevel = ShaftOpeningLevel(opening, BuiltInParameter.WALL_HEIGHT_TYPE);
                    var bottomOffset = ShaftOpeningLength(opening, BuiltInParameter.WALL_BASE_OFFSET);
                    var topOffset = ShaftOpeningLength(opening, BuiltInParameter.WALL_TOP_OFFSET);
                    var height = ShaftOpeningLength(opening, BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    if (bottomLevel != null && bottomOffset.HasValue)
                    {
                        double bottom = bottomLevel.ProjectElevation + bottomOffset.Value;
                        double? top = topLevel != null && topOffset.HasValue ? topLevel.ProjectElevation + topOffset.Value :
                            height.HasValue ? bottom + height.Value : (double?)null;
                        if (top.HasValue && top.Value > bottom + FloorGeometryTolerance)
                            return Tuple.Create(source.Transform.OfPoint(new XYZ(0, 0, bottom)).Z,
                                source.Transform.OfPoint(new XYZ(0, 0, top.Value)).Z);
                    }
                }
                if (bounds != null && bounds[5] > bounds[2] + FloorGeometryTolerance)
                    return Tuple.Create(bounds[2], bounds[5]);
                // A slab's overall Z bounds do not prove that its opening exists throughout
                // that range (sloped/stepped slabs). Its actual section is checked separately.
                return null;
            }

            private ShaftSectionRead ReadShaftOpening(Source source, Opening opening, double elevation, double[] bounds)
            {
                string key = ShaftEvidenceKey(source, opening) + "/" + elevation.ToString("R", CultureInfo.InvariantCulture);
                ShaftSectionRead result;
                if (shaftEvidenceOpenings.TryGetValue(key, out result)) return result;
                result = new ShaftSectionRead { Region = PlanarRegion.Empty };
                bool cache = true;
                try
                {
                    // A wall opening is a vertical face, not a floor void.
                    if (opening.Host is Wall) return result;
                    if (opening.Host != null && !PhaseAccepted(source, opening.Host)) return result;
                    var range = ShaftOpeningRange(source, opening, bounds);
                    if (range != null && (elevation < range.Item1 - FloorGeometryTolerance || elevation > range.Item2 + FloorGeometryTolerance)) return result;
                    var curves = new List<Curve>();
                    if (opening.IsRectBoundary)
                    {
                        var rectangle = opening.BoundaryRect;
                        if (rectangle == null || rectangle.Count != 2)
                            throw new InvalidOperationException("Не прочитана прямоугольная граница проёма.");
                        var a = rectangle[0]; var b = rectangle[1];
                        // BoundaryRect is native boundary data, not Element.get_BoundingBox.
                        // Non-horizontal rectangles cannot establish a shaft footprint.
                        if (Math.Abs(a.Z - b.Z) > FloorGeometryTolerance) return result;
                        if (Math.Abs(a.X - b.X) <= FloorGeometryTolerance || Math.Abs(a.Y - b.Y) <= FloorGeometryTolerance)
                            throw new InvalidOperationException("Горизонтальная граница проёма вырождена.");
                        var points = new[] { a, new XYZ(b.X, a.Y, a.Z), b, new XYZ(a.X, b.Y, a.Z) };
                        for (int i = 0; i < points.Length; i++) curves.Add(Line.CreateBound(points[i], points[(i + 1) % points.Length]));
                    }
                    else
                    {
                        var boundary = opening.BoundaryCurves;
                        if (boundary == null || boundary.Size == 0) throw new InvalidOperationException("Не прочитана граница проёма.");
                        curves.AddRange(boundary.Cast<Curve>());
                    }
                    var transformed = curves.Select(c => c.CreateTransformed(source.Transform)).ToList();
                    var points3d = transformed.SelectMany(c => CurvePoints(c, ArcChordTolerance)).ToList();
                    if (points3d.Count == 0) throw new InvalidOperationException("Граница проёма пуста.");
                    double plane = points3d[0][2];
                    if (points3d.Any(p => Math.Abs(p[2] - plane) > FloorGeometryTolerance)) return result;
                    if (range == null) throw new InvalidOperationException("Не определён диапазон высоты проёма; плоская граница не доказывает его наличие на этом этаже.");
                    result.Region = NativeBoundaryRegion(transformed, plane);
                    if (result.Region.IsEmpty) throw new InvalidOperationException("Граница проёма не образует положительную площадь.");
                }
                catch (OperationCanceledException) { cache = false; throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { cache = false; throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { cache = false; throw; }
                catch (Exception ex) { result.Error = ex.Message; }
                finally { if (cache) shaftEvidenceOpenings[key] = result; }
                return result;
            }

            private void AddShaftMaterialVoids(ShaftPhysicalEvidence evidence, PlanarRegion material, bool complete,
                IEnumerable<string> provenance, string kind, double elevation)
            {
                if (!complete || material.IsEmpty) return;
                try
                {
                    string description = kind + " на отметке " + ShaftEvidenceElevation(elevation) + ": " +
                        string.Join("; ", provenance.Distinct(StringComparer.Ordinal).OrderBy(x => x, StringComparer.Ordinal));
                    var voids = EnclosedShaftVoids(material).Components()
                        .Select(paths => new ShaftVoidEvidence { Region = new PlanarRegion(paths), Description = description }).ToList();
                    evidence.Voids.AddRange(voids);
                }
                catch (OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                catch (Exception ex) { evidence.Errors.Add(kind + ": не прочитаны замкнутые пустоты. " + ex.Message); }
            }

            private ShaftPhysicalEvidence ReadShaftPhysicalEvidence(Record candidate)
            {
                var evidence = new ShaftPhysicalEvidence();
                if (candidate?.AutomaticShaftRegion == null || candidate.AutomaticShaftRegion.IsEmpty || candidate.Source == null || candidate.Level == null)
                    throw new InvalidOperationException("У кандидата шахты отсутствует контур, источник или расчётный уровень.");
                var points = candidate.AutomaticShaftRegion.Paths.SelectMany(p => p).ToList();
                var search = new[] { points.Min(p => p.X) / PlanarRegion.Scale, points.Min(p => p.Y) / PlanarRegion.Scale,
                    points.Max(p => p.X) / PlanarRegion.Scale, points.Max(p => p.Y) / PlanarRegion.Scale };
                var walls = PlanarRegion.Empty; var floors = PlanarRegion.Empty;
                bool floorsComplete = true;
                var wallProvenance = new List<string>(); var floorProvenance = new List<string>();
                foreach (var source in Sources.Where(s => s.Loaded && s.LoadError == null && s.Mode == candidate.Source.Mode &&
                    Math.Abs(s.Transform.BasisZ.DotProduct(XYZ.BasisZ) - 1) < 1e-8))
                {
                    var visited = new HashSet<string>(StringComparer.Ordinal);
                    var physical = source.Elements.Where(e => e is Wall || e is Floor || e is Opening ||
                            e is FamilyInstance && IsPhysicalShellCandidate(e))
                        .SelectMany(e => e is Wall ? WallMembers((Wall)e).Cast<Element>() : new[] { e });
                    foreach (var element in physical)
                    {
                        if (!visited.Add(element.UniqueId) || !PhaseAccepted(source, element)) continue;
                        bool opening = element is Opening;
                        bool vertical = !opening && !(element is Floor);
                        string description = ShaftEvidenceDescription(source, element);
                        var bounds = ReadShaftEvidenceBounds(source, element);
                        if (bounds.Error != null || bounds.Bounds == null)
                        {
                            // A missing broad-phase extent must not silently omit material that could fill a void.
                            if (!opening)
                            {
                                if (vertical) evidence.WallsComplete = false; else floorsComplete = false;
                                evidence.Errors.Add(description + ": не прочитаны габариты конструкции. " + (bounds.Error ?? "Revit не вернул габариты."));
                                continue;
                            }
                            if (bounds.Error != null) { evidence.Errors.Add(description + ": не прочитаны габариты проёма. " + bounds.Error); continue; }
                        }
                        if (bounds.Bounds != null && !ShaftEvidenceBoundsOverlap(bounds.Bounds, search)) continue;
                        if (opening)
                        {
                            var read = ReadShaftOpening(source, (Opening)element, candidate.Z, bounds.Bounds);
                            if (read.Error != null) evidence.Errors.Add(description + ": " + read.Error);
                            else if (read.Region != null && !read.Region.IsEmpty)
                                foreach (var paths in read.Region.Components())
                                    evidence.Voids.Add(new ShaftVoidEvidence { Region = new PlanarRegion(paths), Description = "Нативный проём: " + description });
                            continue;
                        }
                        // Both endpoints are included: a slab's actual upper face at the floor datum is evidence too.
                        if (candidate.Z < bounds.Bounds[2] - FloorGeometryTolerance || candidate.Z > bounds.Bounds[5] + FloorGeometryTolerance) continue;
                        var section = ReadShaftMaterialSection(source, element, candidate.Z);
                        if (section.Error != null)
                        {
                            if (vertical) evidence.WallsComplete = false; else floorsComplete = false;
                            evidence.Errors.Add(description + ": не прочитано сечение на отметке " + ShaftEvidenceElevation(candidate.Z) + ". " + section.Error);
                            continue;
                        }
                        try
                        {
                            if (vertical) { walls = RegionUnion(walls, section.Region); if (!section.Region.IsEmpty) wallProvenance.Add(description); }
                            else { floors = RegionUnion(floors, section.Region); if (!section.Region.IsEmpty) floorProvenance.Add(description); }
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                        catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                        catch (Exception ex)
                        {
                            if (vertical) evidence.WallsComplete = false; else floorsComplete = false;
                            evidence.Errors.Add(description + ": не объединено физическое сечение. " + ex.Message);
                        }
                    }
                }
                evidence.WallMaterial = walls;
                AddShaftMaterialVoids(evidence, walls, evidence.WallsComplete, wallProvenance, "Замкнутая полость между стенами и другими ограждениями", candidate.Z);
                AddShaftMaterialVoids(evidence, floors, floorsComplete, floorProvenance, "Замкнутое отверстие в перекрытиях", candidate.Z);
                return evidence;
            }
        }
    }
}

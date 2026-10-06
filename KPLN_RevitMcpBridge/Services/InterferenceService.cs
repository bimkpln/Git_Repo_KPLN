using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class InterferenceService
    {
        private sealed class GeometryCoverage
        {
            internal bool Solid;
            internal bool Surface;
            internal bool Curve;
            internal int VolumeSolidCount;
            internal int ZeroVolumeSolidCount;
            internal int MeshCount;
            internal int FaceCount;
        }

        private static void InspectGeometry(GeometryElement geometry, GeometryCoverage coverage, int depth)
        {
            if (geometry == null)
                return;

            if (depth > 32)
                throw new InvalidOperationException("Превышена глубина вложенной геометрии.");

            foreach (GeometryObject item in geometry)
            {
                var solid = item as Solid;
                if (solid != null)
                {
                    if (solid.Faces.Size > 0 && Math.Abs(solid.Volume) > 0)
                    {
                        coverage.Solid = true;
                        coverage.VolumeSolidCount++;
                    }
                    else if (solid.Faces.Size > 0)
                    {
                        coverage.Surface = true;
                        coverage.ZeroVolumeSolidCount++;
                    }

                    continue;
                }

                var instance = item as GeometryInstance;
                if (instance != null)
                {
                    using (var nested = instance.GetInstanceGeometry())
                        InspectGeometry(nested, coverage, depth + 1);

                    continue;
                }

                var element = item as GeometryElement;
                if (element != null)
                    InspectGeometry(element, coverage, depth + 1);
                else if (item is Mesh || item is Face)
                {
                    coverage.Surface = true;
                    if (item is Mesh)
                        coverage.MeshCount++;
                    else
                        coverage.FaceCount++;
                }
                else if (item is Curve || item is PolyLine)
                    coverage.Curve = true;
            }
        }

        internal static object Check(Document doc, IDictionary<string, object> input)
        {
            if (doc.IsFamilyDocument)
                throw new BridgeException("not_project_document", "Проверка требует документа проекта.");

            // Старый клиент не должен незаметно ограничить новую проверку категориями/видом.
            foreach (var key in new[] { "left_category_ids", "right_category_ids", "category_id", "view_unique_id", "unique_ids" })
            {
                if (input.Get(key) != null)
                    throw new BridgeException("invalid_input", "check_intersections проверяет весь документ; параметр " + key + " больше не поддерживается.");
            }

            int maxPairs = input.Range("max_pairs", 1000, 1, 5000);
            int seconds = input.Range("time_budget_seconds", 20, 1, 60);
            var clock = Stopwatch.StartNew();
            var solids = new List<Element>();
            var unsupported = new List<object>();
            var errors = new List<object>();
            var links = new List<object>();
            var excluded = new Dictionary<string, int>();
            var nonVolumetric = new List<object>();
            var geometryNotes = new List<object>();
            int scanned = 0;
            string stopReason = null;
            Element[] inventory;

            using (var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType())
                inventory = collector.OrderBy(e => RevitService.IdValue(e.Id)).ToArray();

            Action<string> exclude = reason =>
            {
                if (!excluded.ContainsKey(reason))
                    excluded[reason] = 0;

                excluded[reason]++;
            };


            // Геометрия документа, а не активного 3D-вида. Никакого списка категорий.
            using (var options = new Options { DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = true, ComputeReferences = false })
            {
                foreach (var element in inventory)
                {
                    if (clock.Elapsed.TotalSeconds >= seconds)
                    {
                        stopReason = "geometry_time_budget";
                        break;
                    }

                    scanned++;
                    // Помещения и камеры не являются физическими препятствиями.
                    long categoryId = element.Category == null ? 0 : RevitService.IdValue(element.Category.Id);
                    if (categoryId == (long)BuiltInCategory.OST_Rooms || categoryId == (long)BuiltInCategory.OST_Cameras)
                    {
                        exclude(categoryId == (long)BuiltInCategory.OST_Rooms ? "rooms_excluded_by_policy" : "cameras_excluded_by_policy");
                        continue;
                    }

                    if (element.ViewSpecific || element is View)
                    {
                        exclude("view_specific_or_view");
                        continue;
                    }

                    if (element is Group)
                    {
                        // Члены групп уже находятся в общей выборке экземпляров.
                        exclude("group_container_members_scanned_individually");
                        continue;
                    }

                    if (element is RevitLinkInstance)
                    {
                        links.Add(RevitService.ElementInfo(element));
                        continue;
                    }

                    try
                    {
                        var coverage = new GeometryCoverage();
                        using (var geometry = element.get_Geometry(options))
                            InspectGeometry(geometry, coverage, 0);

                        if (!coverage.Solid)
                        {
                            var box = element.get_BoundingBox(null);
                            if (!coverage.Curve && box != null && box.Max.X > box.Min.X && box.Max.Y > box.Min.Y && box.Max.Z > box.Min.Z)
                            {
                                exclude("no_volume_by_test_model_policy");
                                nonVolumetric.Add(RevitService.ElementInfo(element));
                                continue;
                            }

                            exclude(coverage.Curve ? "curve_only_no_volume" : "no_volume_by_test_model_policy");
                            if (coverage.Curve || coverage.Surface)
                                nonVolumetric.Add(RevitService.ElementInfo(element));

                            continue;
                        }

                        // В тестовой модели проверяем только объёмную часть; поверхности — диагностика.
                        if (coverage.Surface)
                        {
                            geometryNotes.Add(new
                            {
                                element = RevitService.ElementInfo(element), reason = "non_volumetric_geometry_ignored_by_test_model_policy",
                                geometry_details = new
                                {
                                    volume_solid_count = coverage.VolumeSolidCount,
                                    zero_volume_solid_count = coverage.ZeroVolumeSolidCount,
                                    mesh_count = coverage.MeshCount, face_count = coverage.FaceCount,
                                    has_volume_geometry = coverage.Solid
                                }
                            });
                        }

                        if (!coverage.Solid)
                            continue;

                        if (!ElementIntersectsFilter.IsElementSupported(element) || !ElementIntersectsFilter.IsCategorySupported(element))
                        {
                            unsupported.Add(new { element = RevitService.ElementInfo(element), reason = "intersection_filter_not_supported" });
                            continue;
                        }

                        solids.Add(element);
                    }
                    catch (Exception ex)
                    {
                        errors.Add(new { element_unique_id = element.UniqueId, stage = "geometry", message = ex.Message });
                    }
                }
            }


            var pairs = new List<object>();
            int checkedSources = 0;
            if (stopReason == null)
            {
                // Самопары/зеркальные пары исключаются до медленного API-фильтра.
                for (int i = 0; i < solids.Count; i++)
                {
                    if (clock.Elapsed.TotalSeconds >= seconds)
                    {
                        stopReason = "intersection_time_budget";
                        break;
                    }

                    var element = solids[i];
                    try
                    {
                        var remaining = solids.Skip(i + 1).Select(e => e.Id).ToList();
                        if (remaining.Count > 0)
                        {
                            using (var filter = new ElementIntersectsElementFilter(element))
                            using (var collector = new FilteredElementCollector(doc, remaining).WherePasses(filter))
                            {
                                foreach (var other in collector)
                                {
                                    if (pairs.Count >= maxPairs)
                                    {
                                        stopReason = "pair_limit";
                                        break;
                                    }

                                    pairs.Add(new { left = RevitService.ElementInfo(element), right = RevitService.ElementInfo(other) });
                                }
                            }
                        }

                        if (stopReason != null)
                            break;

                        checkedSources++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add(new { element_unique_id = element.UniqueId, stage = "intersection", message = ex.Message });
                    }
                }
            }


            // Закрытые РН и связанные файлы явно учитываются как неполнота.
            var closedWorksets = doc.IsWorkshared
                ? new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets().Where(w => !w.IsOpen)
                    .Select(w => new { id = w.Id.IntegerValue, name = w.Name }).ToArray()
                : new[] { new { id = 0, name = "" } }.Take(0).ToArray();
            bool complete = stopReason == null && unsupported.Count == 0 && errors.Count == 0 && links.Count == 0 && closedWorksets.Length == 0;
            string status = !complete ? "incomplete" : solids.Count == 0 ? "empty_scope" : pairs.Count == 0 ? "no_intersections" : "intersections_found";

            return new
            {
                contract_version = 2,
                method = "revit_api_element_intersects_element_filter", native_dialog_executed = false,
                scope = "all_3d_elements_current_document; rooms_and_cameras_excluded; no_user_category_or_view_filters; linked_documents_not_checked",
                document_instance_count = inventory.Length, scanned_instance_count = scanned,
                geometry_scan_complete = scanned == inventory.Length,
                solid_element_count = solids.Count, checked_source_count = checkedSources,
                complete, status, stop_reason = stopReason, elapsed_ms = clock.ElapsedMilliseconds,
                pair_count = pairs.Count, pairs, unsupported_elements = unsupported, errors,
                excluded_counts = excluded, non_volumetric_elements = nonVolumetric,
                geometry_notes = geometryNotes, geometry_policy = "test_model_volume_only",
                unprocessed_link_instances = links, closed_worksets = closedWorksets,
                message = status == "no_intersections" ? "Пересечения не обнаружены (API-проверка объёмных элементов всего текущего документа)." : null,
                limitations = new[]
                {
                    "Проверка не ограничена категорией, выделением или видом; члены групп проверяются как экземпляры.",
                    "Помещения и камеры исключены по правилам проверки и не делают результат неполным.",
                    "Правило тестовой модели: отсутствие объёмного Solid исключает элемент; поверхности и нулевые тела не делают результат неполным.",
                    "Автоматически соединённые элементы могут не определяться как пересечения самим Revit API.",
                    "При связях, закрытых РН, ошибках геометрии или лимитах результат неполный. Нулевой список пар не означает успех.",
                    "Бюджет проверяется между API-вызовами; отдельный вызов API нельзя безопасно прервать."
                }
            };
        }
    }
}

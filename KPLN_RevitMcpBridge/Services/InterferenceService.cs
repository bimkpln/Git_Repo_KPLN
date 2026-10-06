using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class InterferenceService
    {
        internal static object OpenNative(UIApplication app)
        {
            var command = RevitCommandId.LookupPostableCommandId(PostableCommand.RunInterferenceCheck);
            if (!app.CanPostCommand(command))
                throw new BridgeException("command_unavailable", "Штатная проверка пересечений сейчас недоступна.", 409);

            app.PostCommand(command);
            return new
            {
                status = "posted", command = "RunInterferenceCheck", check_completed = false,
                requires_user_interaction = true, result = (object)null,
                message = "Команда поставлена в очередь Revit. Категории и запуск задаются в штатном окне; это не результат проверки."
            };
        }

        private static ElementId[] Categories(Document doc, IDictionary<string, object> input, string key)
        {
            var ids = input.List(key, 64).Cast<object>().Select(v => RevitService.Id(Json.Integer(v, key))).ToArray();
            if (ids.Distinct().Count() != ids.Length)
                throw new BridgeException("invalid_input", key + ": категории не должны повторяться.");

            foreach (var id in ids)
            {
                var category = Category.GetCategory(doc, id);
                if (category == null || category.CategoryType != CategoryType.Model)
                    throw new BridgeException("invalid_input", key + ": требуется существующая категория модели.");
            }

            return ids;
        }

        private static Element[] Collect(Document doc, ElementId[] categories)
        {
            using (var filter = new ElementMulticategoryFilter(categories))
            using (var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter))
            {
                // Ограничение защищает UI-поток: большие проверки требуют отдельного пакетного сценария.
                var elements = collector.Take(2001).ToArray();
                if (elements.Length > 2000)
                    throw new BridgeException("scope_too_large", "В каждой стороне допускается не более 2000 элементов. Сузьте категории.");

                return elements.OrderBy(e => RevitService.IdValue(e.Id)).ToArray();
            }
        }

        internal static object Check(Document doc, IDictionary<string, object> input)
        {
            if (doc.IsFamilyDocument)
                throw new BridgeException("not_project_document", "Проверка требует документа проекта.");

            var leftCategories = Categories(doc, input, "left_category_ids");
            var rightCategories = Categories(doc, input, "right_category_ids");
            int maxPairs = input.Range("max_pairs", 1000, 1, 5000);
            var clock = Stopwatch.StartNew();
            var left = Collect(doc, leftCategories);
            var right = Collect(doc, rightCategories);
            var unsupported = left.Concat(right).GroupBy(e => e.Id).Select(g => g.First())
                .Where(e => !ElementIntersectsFilter.IsElementSupported(e) || !ElementIntersectsFilter.IsCategorySupported(e)).ToArray();
            var unsupportedIds = new HashSet<ElementId>(unsupported.Select(e => e.Id));
            var rightIds = right.Where(e => !unsupportedIds.Contains(e.Id)).Select(e => e.Id).ToList();
            var pairs = new List<object>();
            var seenPairs = new HashSet<string>();
            var errors = new List<object>();
            int checkedSources = 0;
            string stopReason = null;


            foreach (var element in left)
            {
                if (clock.Elapsed.TotalSeconds >= 20)
                {
                    stopReason = "time_budget";
                    break;
                }

                if (unsupportedIds.Contains(element.Id))
                    continue;

                try
                {
                    if (rightIds.Count > 0)
                    {
                        using (var filter = new ElementIntersectsElementFilter(element))
                        using (var collector = new FilteredElementCollector(doc, rightIds).WherePasses(filter))
                        {
                            foreach (var other in collector)
                            {
                                long a = RevitService.IdValue(element.Id), b = RevitService.IdValue(other.Id);
                                if (a == b || !seenPairs.Add(Math.Min(a, b) + ":" + Math.Max(a, b)))
                                    continue;

                                // Лимит не превращается в ложное «пересечений нет» или полный отчёт.
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
                catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                {
                    errors.Add(new { element_unique_id = element.UniqueId, message = ex.Message });
                }
            }


            bool complete = stopReason == null && unsupported.Length == 0 && errors.Count == 0;
            bool empty = left.Length == 0 || right.Length == 0;
            string status = !complete ? "incomplete" : empty ? "empty_scope" : pairs.Count == 0 ? "no_intersections" : "intersections_found";
            return new
            {
                method = "revit_api_element_intersects_element_filter", native_dialog_executed = false,
                scope = "current_document_only; all non-type elements in explicit category sets; no linked documents",
                left_category_ids = leftCategories.Select(ModelInspectionService.IdText).ToArray(),
                right_category_ids = rightCategories.Select(ModelInspectionService.IdText).ToArray(),
                left_count = left.Length, right_count = right.Length, checked_source_count = checkedSources,
                complete, status, stop_reason = stopReason, elapsed_ms = clock.ElapsedMilliseconds,
                pair_count = pairs.Count, pairs, unsupported_elements = unsupported.Select(RevitService.ElementInfo).ToArray(), errors,
                message = status == "no_intersections" ? "Пересечения не обнаружены (проверка Revit API в указанной области)." : null,
                limitations = new[]
                {
                    "Логика Autodesk Interference Report, но не запуск и не экспорт штатного диалога.",
                    "Не обнаруживает элементы без solid-геометрии и некоторые автоматически соединённые элементы, как штатный механизм.",
                    "Полнота относится к выполнению API-проверки выбранных категорий, а не ко всем возможным геометрическим дефектам.",
                    "Бюджет времени проверяется между вызовами Revit API; отдельный вызов API нельзя безопасно прервать."
                }
            };
        }
    }
}

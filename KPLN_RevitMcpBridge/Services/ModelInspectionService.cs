using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Чтение фактов модели. Критерии оценки остаются во внешнем слое правил.</summary>
    internal static class ModelInspectionService
    {
        internal static string IdText(ElementId id) => RevitService.IdValue(id).ToString(CultureInfo.InvariantCulture);

        private static T Require<T>(Document doc, IDictionary<string, object> input, string key) where T : Element
        {
            var element = doc.GetElement(input.Text(key, true)) as T;
            if (element == null)
                throw new BridgeException("wrong_element_type", key + ": элемент не найден или имеет неподходящий класс.");

            return element;
        }

        private static object Member(Element element)
        {
            return new
            {
                element = RevitService.ElementInfo(element), group_id = IdText(element.GroupId),
                owner_view_id = IdText(element.OwnerViewId), level_id = IdText(element.LevelId),
                workset_id = element.WorksetId.IntegerValue, is_nested_group = element is Group
            };
        }

        internal static object GroupMembers(Document doc, IDictionary<string, object> input)
        {
            var group = Require<Group>(doc, input, "group_unique_id");
            int offset = input.Range("offset", 0, 0, 10000000);
            int limit = input.Range("limit", 100, 1, 200);
            // Порядок GetMemberIds сохраняется: он полезен для сопоставления экземпляров группы.
            var ids = group.GetMemberIds().ToArray();
            var attachedTypes = group.Category != null && RevitService.IdValue(group.Category.Id) == (long)BuiltInCategory.OST_IOSModelGroups
                ? group.GetAvailableAttachedDetailGroupTypeIds().Select(IdText).ToArray() : new string[0];

            return new
            {
                group = RevitService.ElementInfo(group), group_type = RevitService.ElementInfo(group.GroupType),
                parent_group_id = IdText(group.GroupId), owner_view_id = IdText(group.OwnerViewId),
                is_attached = group.IsAttached, attached_parent_id = group.IsAttached ? IdText(group.AttachedParentId) : null,
                available_attached_detail_group_type_ids = attachedTypes,
                membership = "direct_only; query nested groups separately", total_count = ids.Length, offset,
                items = ids.Skip(offset).Take(limit).Select((id, index) => new { member_index = offset + index, data = Member(doc.GetElement(id)) }).ToArray(),
                next_offset = offset + limit < ids.Length ? (int?)(offset + limit) : null
            };
        }

        private static object ViewInfo(View view)
        {
            return new
            {
                element = RevitService.ElementInfo(view), view_type = view.ViewType.ToString(), scale = view.Scale,
                template_id = IdText(view.ViewTemplateId), detail_level = view.DetailLevel.ToString(),
                display_style = view.DisplayStyle.ToString()
            };
        }

        internal static object SheetContents(Document doc, IDictionary<string, object> input)
        {
            var sheet = Require<ViewSheet>(doc, input, "sheet_unique_id");
            using (var schedules = new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance)))
            using (var titles = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType())
            {
                return new
                {
                    sheet = RevitService.ElementInfo(sheet), sheet_number = sheet.SheetNumber, is_placeholder = sheet.IsPlaceholder,
                    viewports = sheet.GetAllViewports().OrderBy(RevitService.IdValue).Select(id => (Viewport)doc.GetElement(id)).Select(vp => new
                    {
                        viewport = RevitService.ElementInfo(vp), view = ViewInfo((View)doc.GetElement(vp.ViewId)),
                        center_mm = new[] { vp.GetBoxCenter().X * 304.8, vp.GetBoxCenter().Y * 304.8 }
                    }).ToArray(),
                    schedules = schedules.Cast<ScheduleSheetInstance>().Where(s => s.OwnerViewId == sheet.Id).OrderBy(s => RevitService.IdValue(s.Id)).Select(s => new
                    {
                        instance = RevitService.ElementInfo(s), schedule_id = IdText(s.ScheduleId),
                        schedule_unique_id = doc.GetElement(s.ScheduleId)?.UniqueId
                    }).ToArray(),
                    title_blocks = titles.Where(t => t.OwnerViewId == sheet.Id).Select(RevitService.ElementInfo).ToArray()
                };
            }
        }

        private static View DrawableView(Document doc, IDictionary<string, object> input)
        {
            var view = Require<View>(doc, input, "view_unique_id");
            if (view.IsTemplate || !FilteredElementCollector.IsViewValidForElementIteration(doc, view.Id))
                throw new BridgeException("unsupported_view", "Вид не поддерживает чтение отображаемых элементов.");

            return view;
        }

        internal static object ViewElements(Document doc, IDictionary<string, object> input)
        {
            var view = DrawableView(doc, input);
            int offset = input.Range("offset", 0, 0, 10000000);
            int limit = input.Range("limit", 100, 1, 200);
            using (var collector = new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType())
            {
                if (input.Get("category_id") != null)
                    collector.OfCategoryId(RevitService.Id(Json.Integer(input.Get("category_id"), "category_id")));

                var elements = collector.OrderBy(e => RevitService.IdValue(e.Id)).Skip(offset).Take(limit + 1).ToArray();
                return new
                {
                    view = ViewInfo(view), visibility_semantics = "potentially_visible_not_rendered; crop and occlusion may exclude elements",
                    offset, items = elements.Take(limit).Select(Member).ToArray(),
                    next_offset = elements.Length > limit ? (int?)(offset + limit) : null
                };
            }
        }

        internal static object ViewVisibility(Document doc, IDictionary<string, object> input)
        {
            var view = DrawableView(doc, input);
            var elements = RevitService.Resolve(doc, input);
            HashSet<ElementId> candidates;
            using (var collector = new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType())
                candidates = new HashSet<ElementId>(collector.ToElementIds());

            var filterIds = view.GetFilters();
            var worksets = doc.IsWorkshared ? new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets() : new List<Workset>();
            var defaults = doc.IsWorkshared ? WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc) : null;

            return new
            {
                view = ViewInfo(view), temporary_hide_isolate_active = view.IsTemporaryHideIsolateActive(),
                visibility_semantics = "potentially_visible_not_rendered; not a pixel-level proof of sheet appearance",
                filters = filterIds.Select(id => new { filter = RevitService.ElementInfo(doc.GetElement(id)), visible = view.GetFilterVisibility(id), enabled = FilterEnabled(view, id) }).ToArray(),
                worksets = worksets.Select(ws => new
                {
                    id = ws.Id.IntegerValue, name = ws.Name, is_open = ws.IsOpen,
                    view_visibility = view.GetWorksetVisibility(ws.Id).ToString(), visible_by_default = defaults.IsWorksetVisible(ws.Id)
                }).ToArray(),
                elements = elements.Select(e => new
                {
                    data = Member(e), potentially_visible = candidates.Contains(e.Id), individually_hidden = e.IsHidden(view),
                    category_hidden = e.Category == null ? (bool?)null : view.GetCategoryHidden(e.Category.Id),
                    workset_visibility = doc.IsWorkshared ? view.GetWorksetVisibility(e.WorksetId).ToString() : null,
                    matched_filters = filterIds.Where(id => MatchesFilter(doc, id, e)).Select(IdText).ToArray(),
                    shown_attached_detail_group_type_ids = e is Group && e.Category != null && RevitService.IdValue(e.Category.Id) == (long)BuiltInCategory.OST_IOSModelGroups
                        ? ((Group)e).GetShownAttachedDetailGroupTypeIds(view).Select(IdText).ToArray() : new string[0]
                }).ToArray()
            };
        }

        private static bool MatchesFilter(Document doc, ElementId id, Element element)
        {
            var filter = doc.GetElement(id);
            var parameterFilter = filter as ParameterFilterElement;
            if (parameterFilter != null)
            {
                if (element.Category == null || !parameterFilter.GetCategories().Contains(element.Category.Id))
                    return false;

                var rule = parameterFilter.GetElementFilter();
                return rule == null || rule.PassesFilter(element);
            }

            var selectionFilter = filter as SelectionFilterElement;
            return selectionFilter != null && selectionFilter.GetElementIds().Contains(element.Id);
        }

        private static bool FilterEnabled(View view, ElementId id)
        {
#if Debug2020 || Revit2020
            // В Revit 2020 нет отдельного выключателя применения фильтра.
            return true;
#else
            return view.GetIsFilterEnabled(id);
#endif
        }

        internal static object ScheduleData(Document doc, IDictionary<string, object> input)
        {
            var schedule = Require<ViewSchedule>(doc, input, "schedule_unique_id");
            var sectionName = input.Text("section") ?? "body";
            if (sectionName != "body" && sectionName != "header")
                throw new BridgeException("invalid_input", "section: ожидается body или header.");

            int offset = input.Range("offset", 0, 0, 10000000);
            int limit = input.Range("limit", 50, 1, 100);
            int columnOffset = input.Range("column_offset", 0, 0, 10000000);
            int columnLimit = input.Range("column_limit", 50, 1, 50);
            var definition = schedule.Definition;
            var sectionType = sectionName == "body" ? SectionType.Body : SectionType.Header;
            var section = schedule.GetTableData().GetSectionData(sectionType);
            var rows = Enumerable.Range(0, section.NumberOfRows).Skip(offset).Take(limit).Select(rowOffset => new
            {
                row_index = section.FirstRowNumber + rowOffset,
                cells = Enumerable.Range(0, section.NumberOfColumns).Skip(columnOffset).Take(columnLimit).Select(colOffset => new
                {
                    column_index = section.FirstColumnNumber + colOffset,
                    text = schedule.GetCellText(sectionType, section.FirstRowNumber + rowOffset, section.FirstColumnNumber + colOffset)
                }).ToArray()
            }).ToArray();

            return new
            {
                schedule = RevitService.ElementInfo(schedule), category_id = IdText(definition.CategoryId),
                is_itemized = definition.IsItemized, show_title = definition.ShowTitle, show_headers = definition.ShowHeaders,
                show_grand_total = definition.ShowGrandTotal, show_grand_total_count = definition.ShowGrandTotalCount,
                show_grand_total_title = definition.ShowGrandTotalTitle, includes_linked_files = definition.IncludeLinkedFiles,
                has_embedded_schedule = definition.HasEmbeddedSchedule,
                definition_scope = "main_definition; embedded definition and calculated expressions are not read",
                fields = definition.GetFieldOrder().Select(id => definition.GetField(id)).Select(f => new
                {
                    field_id = f.FieldId.IntegerValue, parameter_id = IdText(f.ParameterId), name = f.GetName(),
                    column_heading = f.ColumnHeading, field_type = f.FieldType.ToString(), is_hidden = f.IsHidden,
                    is_calculated = f.IsCalculatedField, display_type = f.DisplayType.ToString()
                }).ToArray(),
                sorting_grouping = definition.GetSortGroupFields().Select(f => new
                {
                    field_id = f.FieldId.IntegerValue, sort_order = f.SortOrder.ToString(), show_header = f.ShowHeader,
                    show_footer = f.ShowFooter, show_footer_title = f.ShowFooterTitle, show_footer_count = f.ShowFooterCount,
                    show_blank_line = f.ShowBlankLine
                }).ToArray(),
                filters = definition.GetFilters().Select(f => new { field_id = f.FieldId.IntegerValue, filter_type = f.FilterType.ToString(), value = FilterValue(f) }).ToArray(),
                section = sectionName, row_count = section.NumberOfRows, column_count = section.NumberOfColumns,
                first_row_index = section.FirstRowNumber, first_column_index = section.FirstColumnNumber,
                offset, column_offset = columnOffset, rows,
                next_offset = offset + limit < section.NumberOfRows ? (int?)(offset + limit) : null,
                next_column_offset = columnOffset + columnLimit < section.NumberOfColumns ? (int?)(columnOffset + columnLimit) : null,
                cell_semantics = "formatted_display_text; not raw numeric values or element-to-row mapping"
            };
        }

        private static object FilterValue(ScheduleFilter filter)
        {
            if (filter.IsStringValue) return new { kind = "string", value = (object)filter.GetStringValue() };
            if (filter.IsIntegerValue) return new { kind = "integer", value = (object)filter.GetIntegerValue() };
            if (filter.IsDoubleValue) return new { kind = "double_revit_internal", value = (object)filter.GetDoubleValue() };
            if (filter.IsElementIdValue) return new { kind = "element_id", value = (object)IdText(filter.GetElementIdValue()) };
            return new { kind = "no_value", value = (object)null };
        }
    }
}

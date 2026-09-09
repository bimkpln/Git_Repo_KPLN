using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal static class RevitMcpToolRegistry
    {
        private static readonly List<RevitMcpToolDefinition> Tools = new List<RevitMcpToolDefinition>
        {
            Tool("get_active_view_in_revit", "ViewContext", "Returns the current active Revit view or sheet information."),
            Tool("get_all_elements_shown_in_view", "Visibility", "Returns element ids visible in a view, sheet or schedule with pagination.", PagedProps(new JProperty("viewOrSheetId", Int("View, sheet or schedule element id. Optional.")))),

            Tool("get_category_by_keyword", "Categories", "Finds Revit categories by keyword.", Props(new JProperty("keyword", Str("Category keyword.")), "keyword")),
            Tool("get_elements_by_category", "Categories", "Returns element ids by Revit category id with pagination.", PagedProps(new JProperty("categoryId", Int("Category id.")), "categoryId")),
            Tool("get_model_categories", "Categories", "Returns all model categories."),
            Tool("get_categories_from_elementids", "Categories", "Returns categories for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_object_classes_from_elementids", "Categories", "Returns Revit API class names for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),

            Tool("get_element_types_for_elementids", "Families", "Returns type ids and type names for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_all_elementids_for_specific_type_ids", "Families", "Returns instance element ids for type ids with pagination.", PagedProps(new JProperty("list_typeIds", IntArray("Type ids.")), "list_typeIds")),
            Tool("get_all_used_families_in_model", "Families", "Returns used families in the model with pagination.", PagedProps()),
            Tool("get_all_used_families_of_category", "Families", "Returns used families for a category with pagination.", PagedProps(new JProperty("categoryId", Int("Category id.")), "categoryId")),
            Tool("get_all_used_types_of_a_family", "Families", "Returns all used types of a family with pagination.", PagedProps(new JProperty("familyName", Str("Exact family name.")), "familyName")),
            Tool("get_all_elements_of_specific_families", "Families", "Returns elements for exact family names with pagination.", PagedProps(new JProperty("familyNames", StrArray("Family names.")), "familyNames")),

            Tool("get_parameters_from_elementid", "Parameters", "Returns parameters for one element or type.", Props(new[] { new JProperty("elementId", Int("Element id.")), new JProperty("getIdValuesAsNames", Bool("Resolve ElementId values as names.")) }, "elementId")),
            Tool("get_parameter_value_for_element_ids", "Parameters", "Returns one parameter value for element ids with pagination.", PagedProps(new[] { new JProperty("list_elementIds", IntArray("Element ids.")), new JProperty("idParameter", Int("Parameter id.")), new JProperty("getIdValuesAsNames", Bool("Resolve ElementId values as names.")) }, "list_elementIds", "idParameter")),
            Tool("get_all_additional_properties_from_elementid", "Parameters", "Returns public Revit API properties for one element.", Props(new JProperty("elementId", Int("Element id.")), "elementId")),
            Tool("get_additional_property_for_all_elementids", "Parameters", "Returns one named public Revit API property for element ids with pagination.", PagedProps(new[] { new JProperty("list_elementIds", IntArray("Element ids.")), new JProperty("propertyName", Str("Exact property name.")) }, "list_elementIds", "propertyName")),
            Tool("get_revitlookup_like_properties", "Parameters", "Returns RevitLookup-like public and special API properties for one element.", Props(new[] { new JProperty("elementId", Int("Element id.")), new JProperty("maxValueLength", Int("Maximum string value length. Default 1000.")) }, "elementId")),
            Tool("get_titleblock_family_parameters_description", "Parameters", "Returns title block family parameter description. The optional description argument can be supplied by the caller.", Props(new JProperty("description", Str("Optional description text.")))),

            Tool("get_location_for_element_ids", "Geometry", "Returns location points or curves for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_boundingboxes_for_element_ids", "Geometry", "Returns bounding boxes for element ids with pagination.", PagedProps(new[] { new JProperty("list_elementIds", IntArray("Element ids.")), new JProperty("idSheet", Int("Optional sheet/view id.")) }, "list_elementIds")),
            Tool("get_boundary_lines", "Geometry", "Returns geometry boundary lines for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_room_boundary_lines", "Geometry", "Returns room boundary lines with pagination.", PagedProps(new JProperty("list_roomIds", IntArray("Room element ids. Optional.")))),
            Tool("get_host_id_for_element_ids", "Geometry", "Returns host element ids for hosted elements with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_material_layers_from_types", "Geometry", "Returns compound structure material layers for type ids with pagination.", PagedProps(new JProperty("list_typeIds", IntArray("Type ids.")), "list_typeIds")),
            UiTool("set_view_section_box_to_elements", "Geometry", "Sets the active 3D view section box around elements.", Props(new[] { new JProperty("list_elementIds", IntArray("Element ids.")), new JProperty("marginMM", Num("Margin in millimeters. Default 500.")) }, "list_elementIds")),

            Tool("get_model_file_info", "ModelInfo", "Returns current model file information."),
            Tool("get_all_project_units", "ModelInfo", "Returns project units."),
            Tool("get_all_warnings_in_the_model", "ModelInfo", "Returns all warnings in the current Revit model with pagination.", PagedProps()),

            Tool("get_all_workset_information", "Worksets", "Returns workset information for the current model."),
            Tool("get_worksets_from_elementids", "Worksets", "Returns workset information for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),
            Tool("get_worksharing_information_for_element_ids", "Worksets", "Returns worksharing information for element ids with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Element ids.")), "list_elementIds")),

            Tool("get_user_selection_in_revit", "Selection", "Returns current Revit user selection with pagination.", PagedProps()),
            UiTool("set_user_selection_in_revit", "Selection", "Sets current Revit user selection.", Props(new JProperty("list_elementIds", IntArray("Element ids to select.")), "list_elementIds")),

            Tool("get_graphic_overrides_for_element_ids_in_view", "Visibility", "Returns element graphic overrides in a view with pagination.", PagedProps(new[] { new JProperty("list_elementIds", IntArray("Element ids.")), new JProperty("viewId", Int("View id.")) }, "list_elementIds", "viewId")),
            Tool("get_graphic_filters_applied_to_views", "Visibility", "Returns graphic filters applied to views with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("View element ids.")), "list_elementIds")),
            Tool("get_all_parameter_filters_in_model", "Visibility", "Returns all parameter filters in the model with pagination.", PagedProps()),
            Tool("get_graphic_overrides_view_filters", "Visibility", "Returns graphic overrides for filters in a view with pagination.", PagedProps(new[] { new JProperty("list_filterIds", IntArray("Filter ids.")), new JProperty("viewId", Int("View id.")) }, "list_filterIds", "viewId")),
            Tool("get_category_visibility_overrides_in_view", "Visibility", "Returns category visibility and graphic overrides in a view with pagination.", PagedProps(new JProperty("viewId", Int("View id.")), "viewId")),
            Tool("get_workset_visibility_in_view", "Visibility", "Returns workset visibility in a view with pagination.", PagedProps(new JProperty("viewId", Int("View id.")), "viewId")),
            Tool("get_link_graphics_overrides_in_view", "Visibility", "Returns link graphic overrides in a view with pagination.", PagedProps(new JProperty("viewId", Int("View id.")), "viewId")),
            Tool("get_detailed_link_graphics_overrides_in_view", "Visibility", "Returns detailed link graphic override information in a view with pagination.", PagedProps(new JProperty("viewId", Int("View id.")), "viewId")),
            Tool("get_all_phases_in_model", "Visibility", "Returns all project phases in chronological order from earliest to latest, including the active view phase. Results are paginated.", PagedProps()),
            Tool("get_phase_visibility_settings", "Visibility", "Returns the view phase, discipline, assigned phase filter settings and phase status for supplied elements. Visibility conclusions cover phase rules only. Revit API 2020/2023/2024 does not expose the phase-status graphic overrides from the Phasing dialog, including colors, line settings, patterns, halftone and materials; tell the user to inspect those settings manually in Revit when they may affect the answer. Element results are paginated.", PagedProps(new[] { new JProperty("viewId", Int("View id. Optional; defaults to the active view.")), new JProperty("list_elementIds", IntArray("Optional element ids whose phase status should be evaluated.")), new JProperty("includeAllFilters", Bool("Include every phase filter defined in the model. Default false.")) })),
            Tool("get_if_elements_pass_filter", "Visibility", "Checks whether element ids pass a parameter filter with pagination.", PagedProps(new[] { new JProperty("filterId", Int("Filter id.")), new JProperty("list_elementIds", IntArray("Element ids.")) }, "filterId", "list_elementIds")),

            Tool("get_viewports_and_schedules_on_sheets", "Schedules", "Returns viewports, schedules and other elements placed on sheets with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Sheet element ids.")), "list_elementIds")),
            Tool("get_schedules_info_and_columns", "Schedules", "Returns schedule structure, fields, filters and columns with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Schedule element ids.")), "list_elementIds")),
            Tool("get_schedule_sorting_info", "Schedules", "Returns schedule sorting and grouping information with pagination.", PagedProps(new JProperty("list_elementIds", IntArray("Schedule element ids.")), "list_elementIds")),

            Tool("get_journal_entries_since", "Journal", "Returns Revit journal entries for a date/time range with pagination. When analyzing the result, focus first on user actions and use technical startup, shutdown and add-in events as supporting context.", Props(new[] { new JProperty("dateTime", Str("Start date/time.")), new JProperty("endDateTime", Str("Optional end date/time.")), new JProperty("limit", Int("Maximum page size in characters. Default and maximum 204800.")), new JProperty("offset", Int("Character offset for the next page. Default 0.")) }, "dateTime")),

            Tool("get_revit_links_in_model", "Links", "Returns Revit links in the current model."),
            Tool("get_revit_link_elements", "Links", "Returns elements from a loaded Revit link with pagination.", Props(new[] { new JProperty("linkInstanceId", Int("RevitLinkInstance id.")), new JProperty("limit", Int("Page size. Default and maximum 200.")), new JProperty("offset", Int("Offset. Default 0.")) }, "linkInstanceId")),
            Tool("get_revit_link_categories", "Links", "Returns categories inside a loaded Revit link with pagination.", PagedProps(new JProperty("linkInstanceId", Int("RevitLinkInstance id.")), "linkInstanceId")),
            Tool("get_revit_link_elements_by_category", "Links", "Returns elements from a loaded Revit link filtered by category with pagination.", Props(new[] { new JProperty("linkInstanceId", Int("RevitLinkInstance id.")), new JProperty("categoryId", Int("Linked category id.")), new JProperty("categoryName", Str("Linked category name fallback.")), new JProperty("limit", Int("Page size. Default and maximum 200.")), new JProperty("offset", Int("Offset. Default 0.")) }, "linkInstanceId")),
            UiTool("get_selected_revit_link_element_id", "Links", "Prompts the user to select linked elements in Revit and returns their ids."),
            Tool("get_revit_link_element_properties", "Links", "Returns properties for an element inside a loaded Revit link.", Props(new[] { new JProperty("linkInstanceId", Int("RevitLinkInstance id.")), new JProperty("linkedElementId", Int("Element id inside linked document.")), new JProperty("getIdValuesAsNames", Bool("Resolve ElementId values as names.")), new JProperty("maxValueLength", Int("Maximum string value length. Default 1000.")), new JProperty("parameterId", Int("Optional parameter id.")), new JProperty("additionalPropertyName", Str("Optional public property name.")) }, "linkInstanceId", "linkedElementId"))
        };

        private static readonly Dictionary<string, RevitMcpToolDefinition> ToolsByName =
            Tools.ToDictionary(i => i.Name, i => i);

        public static IReadOnlyList<RevitMcpToolDefinition> GetAll()
        {
            return Tools;
        }

        public static bool TryGet(string toolName, out RevitMcpToolDefinition definition)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                definition = null;
                return false;
            }

            return ToolsByName.TryGetValue(toolName, out definition);
        }

        public static JArray ToMcpToolsArray()
        {
            return new JArray(Tools.Select(i => i.ToMcpToolObject()));
        }

        private static RevitMcpToolDefinition Tool(string name, string area, string description)
        {
            return Tool(name, area, description, RevitMcpToolDefinition.CreateEmptyInputSchema());
        }

        private static RevitMcpToolDefinition Tool(string name, string area, string description, JObject inputSchema)
        {
            return new RevitMcpToolDefinition
            {
                Name = name,
                Area = area,
                Description = description,
                RiskLevel = RevitMcpToolRiskLevel.ReadOnly,
                InputSchema = inputSchema
            };
        }

        private static RevitMcpToolDefinition UiTool(string name, string area, string description)
        {
            return UiTool(name, area, description, RevitMcpToolDefinition.CreateEmptyInputSchema());
        }

        private static RevitMcpToolDefinition UiTool(string name, string area, string description, JObject inputSchema)
        {
            return new RevitMcpToolDefinition
            {
                Name = name,
                Area = area,
                Description = description,
                RiskLevel = RevitMcpToolRiskLevel.UiChanging,
                InputSchema = inputSchema
            };
        }

        private static JObject Props(JProperty property, params string[] required)
        {
            return Props(new[] { property }, required);
        }

        private static JObject Props(IEnumerable<JProperty> properties, params string[] required)
        {
            JObject schemaProperties = new JObject();
            foreach (JProperty property in properties)
                schemaProperties.Add(property);

            return new JObject
            {
                { "type", "object" },
                { "properties", schemaProperties },
                { "required", new JArray(required ?? new string[0]) }
            };
        }

        private static JObject PagedProps(params string[] required)
        {
            return PagedProps(new JProperty[0], required);
        }

        private static JObject PagedProps(JProperty property, params string[] required)
        {
            return PagedProps(new[] { property }, required);
        }

        private static JObject PagedProps(IEnumerable<JProperty> properties, params string[] required)
        {
            List<JProperty> allProperties = new List<JProperty>();
            if (properties != null)
                allProperties.AddRange(properties);

            allProperties.Add(new JProperty("limit", Int("Page size. Default and maximum 200.")));
            allProperties.Add(new JProperty("offset", Int("Offset. Default 0.")));
            return Props(allProperties, required);
        }

        private static JObject Str(string description)
        {
            return new JObject { { "type", "string" }, { "description", description } };
        }

        private static JObject Int(string description)
        {
            return new JObject { { "type", "integer" }, { "description", description } };
        }

        private static JObject Num(string description)
        {
            return new JObject { { "type", "number" }, { "description", description } };
        }

        private static JObject Bool(string description)
        {
            return new JObject { { "type", "boolean" }, { "description", description } };
        }

        private static JObject IntArray(string description)
        {
            return new JObject
            {
                { "type", "array" },
                { "description", description },
                { "items", new JObject { { "type", "integer" } } }
            };
        }

        private static JObject StrArray(string description)
        {
            return new JObject
            {
                { "type", "array" },
                { "description", description },
                { "items", new JObject { { "type", "string" } } }
            };
        }
    }
}

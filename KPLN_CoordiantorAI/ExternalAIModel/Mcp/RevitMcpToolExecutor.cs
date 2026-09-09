using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.Common;
using KPLN_CoordiantorAI.ExternalModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal class RevitMcpToolExecutor
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;

        public RevitMcpToolExecutor(Document document, UIDocument uiDocument)
        {
            _document = document;
            _uiDocument = uiDocument;
        }

        public RevitMcpToolCallResponse Execute(string toolName, JObject arguments)
        {
            if (!RevitMcpToolRegistry.TryGet(toolName, out RevitMcpToolDefinition definition))
            {
                return CreateErrorResponse(
                    toolName,
                    "tool_not_found",
                    "MCP tool is not registered in Revit MCP registry: " + (toolName ?? "<null>"),
                    null,
                    new JObject
                    {
                        { "toolName", toolName ?? string.Empty },
                        { "availableTools", new JArray(RevitMcpToolRegistry.GetAll().Select(i => i.Name)) }
                    });
            }

            if (_document == null)
            {
                return CreateErrorResponse(
                    definition.Name,
                    "active_document_missing",
                    "Active Revit document is missing. Open a Revit model and retry the MCP tool call.",
                    null,
                    new JObject { { "toolName", definition.Name } });
            }

            if (RequiresUiDocument(definition.Name) && _uiDocument == null)
            {
                return CreateErrorResponse(
                    definition.Name,
                    "ui_document_missing",
                    "Active Revit UI document is missing. Open a model view in Revit and retry the MCP tool call.",
                    null,
                    new JObject { { "toolName", definition.Name } });
            }

            try
            {
                JObject safeArguments = arguments ?? new JObject();
                ValidateArguments(definition, safeArguments);
                object rawResult = ExecuteRegisteredTool(definition.Name, safeArguments);
                JToken result = ToJToken(rawResult);

                RevitMcpToolCallResponse toolResultError = TryCreateToolResultErrorResponse(definition.Name, result);
                if (toolResultError != null)
                    return toolResultError;

                result = ApplyPaginationIfNeeded(definition.Name, result, safeArguments);

                return new RevitMcpToolCallResponse
                {
                    Success = true,
                    ToolName = definition.Name,
                    Result = result
                };
            }
            catch (RevitMcpToolArgumentException ex)
            {
                return CreateErrorResponse(
                    definition.Name,
                    "invalid_arguments",
                    ex.Message,
                    ex.GetType().FullName,
                    ex.Details);
            }
            catch (NotSupportedException ex)
            {
                return CreateErrorResponse(
                    definition.Name,
                    "tool_not_implemented",
                    ex.Message,
                    ex.GetType().FullName,
                    new JObject { { "toolName", definition.Name } });
            }
            catch (Exception ex)
            {
                return CreateErrorResponse(
                    definition.Name,
                    "revit_api_error",
                    string.IsNullOrWhiteSpace(ex.Message) ? "Revit API tool execution failed." : ex.Message,
                    ex.GetType().FullName,
                    new JObject { { "toolName", definition.Name } });
            }
        }
        private object ExecuteRegisteredTool(string toolName, JObject arguments)
        {
            switch (toolName)
            {
                case "get_active_view_in_revit":
                    return Commands.GetActiveViewInfo(_document);

                case "get_model_categories":
                    return Commands.GetModelCategories(_document);

                case "get_user_selection_in_revit":
                    return Commands.GetUserSelectionInRevit(
                        _document,
                        _uiDocument,
                        GetInt(arguments, "limit", 200),
                        GetInt(arguments, "offset"));

                case "get_model_file_info":
                    return Commands.GetModelFileInfo(_document);

                case "get_all_project_units":
                    return Commands.GetAllProjectUnits(_document);

                case "get_category_by_keyword":
                    return Commands.GetCategoryByKeyword(
                        _document,
                        GetString(arguments, "keyword"));

                case "get_elements_by_category":
                    return Commands.GetElementsByCategory(
                        _document,
                        GetInt(arguments, "categoryId"));

                case "get_categories_from_elementids":
                    return Commands.GetCategoriesFromElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_element_types_for_elementids":
                    return Commands.GetElementTypesForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_all_elementids_for_specific_type_ids":
                    return Commands.GetAllElementIdsForSpecificTypeIds(
                        _document,
                        GetIntList(arguments, "list_typeIds"));

                case "get_all_used_families_in_model":
                    return Commands.GetAllUsedFamiliesInModel(_document);

                case "get_all_used_families_of_category":
                    return Commands.GetAllUsedFamiliesOfCategory(
                        _document,
                        GetInt(arguments, "categoryId"));

                case "get_all_used_types_of_a_family":
                    return Commands.GetAllUsedTypesOfAFamily(
                        _document,
                        GetString(arguments, "familyName"));

                case "get_all_elements_of_specific_families":
                    return Commands.GetAllElementsOfSpecificFamilies(
                        _document,
                        GetStringList(arguments, "familyNames"));

                case "get_parameters_from_elementid":
                    return Commands.GetParametersFromElementId(
                        _document,
                        GetInt(arguments, "elementId"),
                        GetBool(arguments, "getIdValuesAsNames"));

                case "get_parameter_value_for_element_ids":
                    return Commands.GetParameterValueForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"),
                        GetInt(arguments, "idParameter"),
                        GetBool(arguments, "getIdValuesAsNames"));

                case "get_all_elements_shown_in_view":
                    return Commands.GetAllElementsShownInView(
                        _document,
                        GetInt(arguments, "viewOrSheetId", IDHelper.ElIdInt(_document.ActiveView.Id)));

                case "get_all_additional_properties_from_elementid":
                    return Commands.GetAllAdditionalPropertiesFromElementId(
                        _document,
                        GetInt(arguments, "elementId"));

                case "get_additional_property_for_all_elementids":
                    return Commands.GetAdditionalPropertyForAllElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"),
                        GetString(arguments, "propertyName"));

                case "get_revitlookup_like_properties":
                    return Commands.GetRevitLookupLikeProperties(
                        _document,
                        GetInt(arguments, "elementId"),
                        GetInt(arguments, "maxValueLength", 1000));

                case "get_location_for_element_ids":
                    return Commands.GetLocationForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_boundingboxes_for_element_ids":
                    return Commands.GetBoundingBoxesForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"),
                        GetNullableInt(arguments, "idSheet"));

                case "get_boundary_lines":
                    return Commands.GetBoundaryLines(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_room_boundary_lines":
                    return Commands.GetRoomBoundaryLines(
                        _document,
                        GetIntList(arguments, "list_roomIds"));

                case "get_host_id_for_element_ids":
                    return Commands.GetHostIdForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_object_classes_from_elementids":
                    return Commands.GetObjectClassesFromElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_material_layers_from_types":
                    return Commands.GetMaterialLayersFromTypes(
                        _document,
                        GetIntList(arguments, "list_typeIds"));

                case "get_all_warnings_in_the_model":
                    return Commands.GetAllWarningsInTheModel(_document);

                case "get_all_workset_information":
                    return Commands.GetAllWorksetInformation(_document);

                case "get_worksets_from_elementids":
                    return Commands.GetWorksetsFromElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_worksharing_information_for_element_ids":
                    return Commands.GetWorksharingInformationForElementIds(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "set_user_selection_in_revit":
                    return Commands.SetUserSelectionInRevit(
                        _document,
                        _uiDocument,
                        GetIntList(arguments, "list_elementIds"));

                case "get_graphic_overrides_for_element_ids_in_view":
                    return Commands.GetGraphicOverridesForElementIdsInView(
                        _document,
                        GetIntList(arguments, "list_elementIds"),
                        GetInt(arguments, "viewId"));

                case "get_graphic_filters_applied_to_views":
                    return Commands.GetGraphicFiltersAppliedToViews(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_all_parameter_filters_in_model":
                    return Commands.GetAllParameterFiltersInModel(_document);

                case "get_graphic_overrides_view_filters":
                    return Commands.GetGraphicOverridesViewFilters(
                        _document,
                        GetIntList(arguments, "list_filterIds"),
                        GetInt(arguments, "viewId"));

                case "get_category_visibility_overrides_in_view":
                    return Commands.GetCategoryVisibilityOverridesInView(
                        _document,
                        GetInt(arguments, "viewId"));

                case "get_workset_visibility_in_view":
                    return Commands.GetWorksetVisibilityInView(
                        _document,
                        GetInt(arguments, "viewId"));

                case "get_link_graphics_overrides_in_view":
                    return Commands.GetLinkGraphicsOverridesInView(
                        _document,
                        GetInt(arguments, "viewId"));

                case "get_detailed_link_graphics_overrides_in_view":
                    return Commands.GetDetailedLinkGraphicsOverridesInView(
                        _document,
                        GetInt(arguments, "viewId"));

                case "get_all_phases_in_model":
                    return Commands.GetAllPhasesInModel(_document);

                case "get_phase_visibility_settings":
                    return Commands.GetPhaseVisibilitySettings(
                        _document,
                        GetNullableInt(arguments, "viewId"),
                        GetIntList(arguments, "list_elementIds"),
                        GetBool(arguments, "includeAllFilters"));

                case "get_viewports_and_schedules_on_sheets":
                    return Commands.GetViewportsAndSchedulesOnSheets(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_schedules_info_and_columns":
                    return Commands.GetSchedulesInfoAndColumns(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_schedule_sorting_info":
                    return Commands.GetScheduleSortingInfo(
                        _document,
                        GetIntList(arguments, "list_elementIds"));

                case "get_if_elements_pass_filter":
                    return Commands.GetIfElementsPassFilter(
                        _document,
                        GetInt(arguments, "filterId"),
                        GetIntList(arguments, "list_elementIds"));

                case "set_view_section_box_to_elements":
                    return Commands.SetViewSectionBoxToElements(
                        _document,
                        _uiDocument,
                        GetIntList(arguments, "list_elementIds"),
                        GetDouble(arguments, "marginMM", 500));

                case "get_journal_entries_since":
                    return Commands.GetJournalEntriesSince(
                        _document,
                        GetString(arguments, "dateTime"),
                        GetString(arguments, "endDateTime", null),
                        GetInt(arguments, "limit", 200 * 1024),
                        GetInt(arguments, "offset"));

                case "get_revit_links_in_model":
                    return Commands.GetRevitLinksInModel(_document);

                case "get_revit_link_elements":
                    return Commands.GetRevitLinkElements(
                        _document,
                        GetInt(arguments, "linkInstanceId"),
                        GetInt(arguments, "limit", 200),
                        GetInt(arguments, "offset"));

                case "get_revit_link_categories":
                    return Commands.GetRevitLinkCategories(
                        _document,
                        GetInt(arguments, "linkInstanceId"));

                case "get_revit_link_elements_by_category":
                    return Commands.GetRevitLinkElementsByCategory(
                        _document,
                        GetInt(arguments, "linkInstanceId"),
                        GetInt(arguments, "categoryId"),
                        GetString(arguments, "categoryName", null),
                        GetInt(arguments, "limit", 200),
                        GetInt(arguments, "offset"));

                case "get_selected_revit_link_element_id":
                    return Commands.GetSelectedRevitLinkElementId(
                        _document,
                        _uiDocument);

                case "get_revit_link_element_properties":
                    return Commands.GetRevitLinkElementProperties(
                        _document,
                        GetInt(arguments, "linkInstanceId"),
                        GetInt(arguments, "linkedElementId"),
                        GetBool(arguments, "getIdValuesAsNames"),
                        GetInt(arguments, "maxValueLength", 1000),
                        GetInt(arguments, "parameterId"),
                        GetString(arguments, "additionalPropertyName", null));

                case "get_titleblock_family_parameters_description":
                    return Commands.GetTitleBlockFamilyParametersDescription(
                        GetString(arguments, "description"));

                default:
                    throw new NotSupportedException("MCP tool is registered but is not implemented in RevitMcpToolExecutor: " + toolName);
            }
        }

        private static int GetInt(JObject arguments, string propertyName, int defaultValue = 0)
        {
            if (arguments == null || arguments[propertyName] == null)
                return defaultValue;

            try
            {
                return arguments[propertyName].Value<int>();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "integer", arguments[propertyName], ex);
            }
        }

        private static string GetString(JObject arguments, string propertyName, string defaultValue = "")
        {
            if (arguments == null || arguments[propertyName] == null)
                return defaultValue;

            try
            {
                return arguments[propertyName].Value<string>() ?? defaultValue;
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "string", arguments[propertyName], ex);
            }
        }

        private static bool GetBool(JObject arguments, string propertyName, bool defaultValue = false)
        {
            if (arguments == null || arguments[propertyName] == null)
                return defaultValue;

            try
            {
                return arguments[propertyName].Value<bool>();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "boolean", arguments[propertyName], ex);
            }
        }

        private static int? GetNullableInt(JObject arguments, string propertyName)
        {
            if (arguments == null || arguments[propertyName] == null || arguments[propertyName].Type == JTokenType.Null)
                return null;

            try
            {
                return arguments[propertyName].Value<int>();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "integer", arguments[propertyName], ex);
            }
        }

        private static double GetDouble(JObject arguments, string propertyName, double defaultValue = 0)
        {
            if (arguments == null || arguments[propertyName] == null)
                return defaultValue;

            try
            {
                return arguments[propertyName].Value<double>();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "number", arguments[propertyName], ex);
            }
        }

        private static List<int> GetIntList(JObject arguments, string propertyName)
        {
            JArray array = arguments == null ? null : arguments[propertyName] as JArray;
            if (array == null)
                return new List<int>();

            try
            {
                return array.Select(i => i.Value<int>()).ToList();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "array of integers", arguments[propertyName], ex);
            }
        }

        private static List<string> GetStringList(JObject arguments, string propertyName)
        {
            JArray array = arguments == null ? null : arguments[propertyName] as JArray;
            if (array == null)
                return new List<string>();

            try
            {
            return array.Select(i => i.Value<string>()).Where(i => !string.IsNullOrWhiteSpace(i)).ToList();
            }
            catch (Exception ex)
            {
                throw CreateInvalidArgumentException(propertyName, "array of strings", arguments[propertyName], ex);
            }
        }

        private static JToken ApplyPaginationIfNeeded(string toolName, JToken result, JObject arguments)
        {
            if (string.Equals(toolName, "get_journal_entries_since", StringComparison.OrdinalIgnoreCase))
                return AddItemsAliasIfNeeded(result, "entries");

            if (string.Equals(toolName, "get_revit_link_elements", StringComparison.OrdinalIgnoreCase)
                || string.Equals(toolName, "get_revit_link_elements_by_category", StringComparison.OrdinalIgnoreCase))
            {
                return AddItemsAliasIfNeeded(result, "elements");
            }

            string collectionProperty = GetPaginatedCollectionProperty(toolName);
            if (collectionProperty == null)
                return result;

            int limit = NormalizeItemLimit(GetInt(arguments, "limit", 200));
            int offset = NormalizeOffset(GetInt(arguments, "offset"));

            if (string.IsNullOrEmpty(collectionProperty))
            {
                if (result is JArray rootArray)
                    return PaginateRootArray(rootArray, limit, offset, "items", "element_ids");

                if (result is JObject rootObject)
                    return PaginateRootObject(rootObject, limit, offset);

                return result;
            }

            JObject resultObject = result as JObject;
            if (resultObject == null)
                return result;

            JToken collection = resultObject[collectionProperty];
            if (collection is JArray array)
                return PaginateArrayProperty(resultObject, collectionProperty, array, limit, offset);

            if (collection is JObject obj)
            {
                if (HasArrayValues(obj))
                    return PaginateObjectOfArraysProperty(resultObject, collectionProperty, obj, limit, offset);

                return PaginateObjectProperty(resultObject, collectionProperty, obj, limit, offset);
            }

            return result;
        }

        private static string GetPaginatedCollectionProperty(string toolName)
        {
            switch (toolName)
            {
                case "get_all_elements_shown_in_view":
                    return "element_ids";
                case "get_elements_by_category":
                    return string.Empty;
                case "get_categories_from_elementids":
                    return string.Empty;
                case "get_object_classes_from_elementids":
                    return "object_classes";
                case "get_element_types_for_elementids":
                    return "type_ids";
                case "get_all_elementids_for_specific_type_ids":
                    return "element_ids_per_type";
                case "get_all_used_families_in_model":
                case "get_all_used_families_of_category":
                    return "families";
                case "get_all_used_types_of_a_family":
                    return "types";
                case "get_all_elements_of_specific_families":
                    return "elements_per_family";
                case "get_parameter_value_for_element_ids":
                    return "parameter_values";
                case "get_additional_property_for_all_elementids":
                    return "property_values";
                case "get_location_for_element_ids":
                    return "locations";
                case "get_boundingboxes_for_element_ids":
                    return "bounding_boxes";
                case "get_boundary_lines":
                case "get_room_boundary_lines":
                    return "boundaries";
                case "get_host_id_for_element_ids":
                    return "host_ids";
                case "get_material_layers_from_types":
                    return "material_layers";
                case "get_all_warnings_in_the_model":
                    return "warnings";
                case "get_worksets_from_elementids":
                    return "workset_assignments";
                case "get_worksharing_information_for_element_ids":
                    return "worksharing_info";
                case "get_graphic_overrides_for_element_ids_in_view":
                    return "overrides";
                case "get_graphic_filters_applied_to_views":
                    return "view_filters";
                case "get_all_parameter_filters_in_model":
                    return "filters";
                case "get_graphic_overrides_view_filters":
                    return "filter_overrides";
                case "get_category_visibility_overrides_in_view":
                    return "categories_overrides";
                case "get_workset_visibility_in_view":
                    return "workset_visibility";
                case "get_link_graphics_overrides_in_view":
                case "get_detailed_link_graphics_overrides_in_view":
                    return "link_overrides";
                case "get_all_phases_in_model":
                    return "phases";
                case "get_phase_visibility_settings":
                    return "elements";
                case "get_viewports_and_schedules_on_sheets":
                    return "sheet_contents";
                case "get_schedules_info_and_columns":
                    return "schedules_info";
                case "get_schedule_sorting_info":
                    return "schedules_sorting";
                case "get_if_elements_pass_filter":
                    return "filter_results";
                case "get_revit_link_categories":
                    return "categories";
                default:
                    return null;
            }
        }

        private static JObject PaginateRootArray(JArray array, int limit, int offset, string itemsProperty, string legacyProperty)
        {
            array = array ?? new JArray();
            JArray page = SliceArray(array, offset, limit);
            JObject result = CreatePaginationMetadata(array.Count, limit, offset, page.Count);
            result[itemsProperty] = page.DeepClone();
            if (!string.IsNullOrEmpty(legacyProperty))
                result[legacyProperty] = page;

            result["count"] = page.Count;
            return result;
        }

        private static JObject PaginateRootObject(JObject obj, int limit, int offset)
        {
            List<JProperty> properties = obj.Properties().ToList();
            List<JProperty> pageProperties = properties.Skip(offset).Take(limit).ToList();

            JObject pageObject = new JObject();
            JArray items = new JArray();
            foreach (JProperty property in pageProperties)
            {
                JToken value = property.Value == null ? JValue.CreateNull() : property.Value.DeepClone();
                pageObject[property.Name] = value;
                items.Add(new JObject
                {
                    { "key", property.Name },
                    { "value", value.DeepClone() }
                });
            }

            JObject result = CreatePaginationMetadata(properties.Count, limit, offset, pageProperties.Count);
            result["items"] = items;
            result["values"] = pageObject;
            result["count"] = pageProperties.Count;
            return result;
        }

        private static JObject PaginateArrayProperty(JObject resultObject, string collectionProperty, JArray array, int limit, int offset)
        {
            JArray page = SliceArray(array, offset, limit);
            JObject result = CloneWithoutPaginationFields(resultObject);
            result[collectionProperty] = page;
            result["items"] = page.DeepClone();
            AddPaginationMetadata(result, array.Count, limit, offset, page.Count);
            SetPageCount(result, page.Count);
            return result;
        }

        private static JObject PaginateObjectOfArraysProperty(JObject resultObject, string collectionProperty, JObject obj, int limit, int offset)
        {
            JObject pageObject = new JObject();
            JArray items = new JArray();
            int totalCount = 0;
            int skipped = 0;
            int taken = 0;

            foreach (JProperty property in obj.Properties())
            {
                JArray array = property.Value as JArray;
                if (array == null)
                {
                    totalCount++;
                    if (skipped++ < offset || taken >= limit)
                        continue;

                    JToken value = property.Value == null ? JValue.CreateNull() : property.Value.DeepClone();
                    pageObject[property.Name] = value;
                    items.Add(new JObject
                    {
                        { "key", property.Name },
                        { "value", value.DeepClone() }
                    });
                    taken++;
                    continue;
                }

                totalCount += array.Count;
                foreach (JToken item in array)
                {
                    if (skipped++ < offset || taken >= limit)
                        continue;

                    if (pageObject[property.Name] == null)
                        pageObject[property.Name] = new JArray();

                    JArray pageArray = pageObject[property.Name] as JArray;
                    JToken value = item == null ? JValue.CreateNull() : item.DeepClone();
                    pageArray.Add(value);
                    items.Add(new JObject
                    {
                        { "key", property.Name },
                        { "value", value.DeepClone() }
                    });
                    taken++;
                }
            }

            JObject result = CloneWithoutPaginationFields(resultObject);
            result[collectionProperty] = pageObject;
            result["items"] = items;
            AddPaginationMetadata(result, totalCount, limit, offset, taken);
            SetPageCount(result, taken);
            return result;
        }

        private static JObject PaginateObjectProperty(JObject resultObject, string collectionProperty, JObject obj, int limit, int offset)
        {
            List<JProperty> properties = obj.Properties().ToList();
            List<JProperty> pageProperties = properties.Skip(offset).Take(limit).ToList();

            JObject pageObject = new JObject();
            JArray items = new JArray();
            foreach (JProperty property in pageProperties)
            {
                JToken value = property.Value == null ? JValue.CreateNull() : property.Value.DeepClone();
                pageObject[property.Name] = value;
                items.Add(new JObject
                {
                    { "key", property.Name },
                    { "value", value.DeepClone() }
                });
            }

            JObject result = CloneWithoutPaginationFields(resultObject);
            result[collectionProperty] = pageObject;
            result["items"] = items;
            AddPaginationMetadata(result, properties.Count, limit, offset, pageProperties.Count);
            SetPageCount(result, pageProperties.Count);
            return result;
        }

        private static bool HasArrayValues(JObject obj)
        {
            if (obj == null)
                return false;

            return obj.Properties().Any(i => i.Value is JArray);
        }

        private static JToken AddItemsAliasIfNeeded(JToken result, string collectionProperty)
        {
            JObject resultObject = result as JObject;
            if (resultObject == null || resultObject["items"] != null)
                return result;

            JToken collection = resultObject[collectionProperty];
            if (collection != null)
                resultObject["items"] = collection.DeepClone();

            bool hasMore = resultObject["has_more"] != null
                && resultObject["has_more"].Type == JTokenType.Boolean
                && resultObject["has_more"].Value<bool>();
            if (hasMore && resultObject["next_offset"] == null)
            {
                int offset = resultObject["offset"] == null ? 0 : resultObject["offset"].Value<int>();
                int returnedCount = resultObject["returned_count"] == null ? 0 : resultObject["returned_count"].Value<int>();
                resultObject["next_offset"] = offset + returnedCount;
            }
            else if (!hasMore && resultObject["next_offset"] == null)
            {
                resultObject["next_offset"] = JValue.CreateNull();
            }

            return resultObject;
        }

        private static JArray SliceArray(JArray array, int offset, int limit)
        {
            JArray page = new JArray();
            foreach (JToken item in array.Skip(offset).Take(limit))
                page.Add(item == null ? JValue.CreateNull() : item.DeepClone());

            return page;
        }

        private static JObject CreatePaginationMetadata(int totalCount, int limit, int offset, int pageCount)
        {
            JObject result = new JObject();
            AddPaginationMetadata(result, totalCount, limit, offset, pageCount);
            return result;
        }

        private static void AddPaginationMetadata(JObject result, int totalCount, int limit, int offset, int pageCount)
        {
            bool hasMore = offset + pageCount < totalCount;
            result["total_count"] = totalCount;
            result["limit"] = limit;
            result["offset"] = offset;
            result["has_more"] = hasMore;
            result["next_offset"] = hasMore ? (JToken)(offset + pageCount) : JValue.CreateNull();
        }

        private static JObject CloneWithoutPaginationFields(JObject source)
        {
            JObject clone = source == null ? new JObject() : (JObject)source.DeepClone();
            clone.Remove("items");
            clone.Remove("total_count");
            clone.Remove("limit");
            clone.Remove("offset");
            clone.Remove("has_more");
            clone.Remove("next_offset");
            return clone;
        }

        private static void SetPageCount(JObject result, int pageCount)
        {
            if (result["count"] != null)
                result["count"] = pageCount;
        }

        private static int NormalizeItemLimit(int limit)
        {
            if (limit <= 0)
                return 200;

            return Math.Min(limit, 200);
        }

        private static int NormalizeOffset(int offset)
        {
            return offset < 0 ? 0 : offset;
        }

        private static bool RequiresUiDocument(string toolName)
        {
            switch (toolName)
            {
                case "get_user_selection_in_revit":
                case "set_user_selection_in_revit":
                case "set_view_section_box_to_elements":
                case "get_selected_revit_link_element_id":
                    return true;
                default:
                    return false;
            }
        }

        private static void ValidateArguments(RevitMcpToolDefinition definition, JObject arguments)
        {
            if (definition == null)
                return;

            JObject schema = definition.InputSchema;
            if (schema == null)
                return;

            JArray required = schema["required"] as JArray;
            if (required != null)
            {
                foreach (JToken requiredItem in required)
                {
                    string propertyName = requiredItem == null ? null : requiredItem.ToString();
                    if (string.IsNullOrWhiteSpace(propertyName))
                        continue;

                    if (arguments == null || arguments[propertyName] == null || arguments[propertyName].Type == JTokenType.Null)
                    {
                        throw new RevitMcpToolArgumentException(
                            "Required MCP tool argument is missing: " + propertyName,
                            new JObject
                            {
                                { "toolName", definition.Name },
                                { "argumentName", propertyName },
                                { "expected", "required" }
                            });
                    }
                }
            }

            JObject properties = schema["properties"] as JObject;
            if (properties == null || arguments == null)
                return;

            foreach (JProperty property in properties.Properties())
            {
                JToken value = arguments[property.Name];
                if (value == null || value.Type == JTokenType.Null)
                    continue;

                JObject propertySchema = property.Value as JObject;
                if (propertySchema == null)
                    continue;

                string expectedType = propertySchema["type"] == null ? null : propertySchema["type"].ToString();
                if (!IsJsonTypeCompatible(value, expectedType))
                {
                    throw new RevitMcpToolArgumentException(
                        "Invalid MCP tool argument type for '" + property.Name + "'. Expected " + expectedType + ", got " + value.Type + ".",
                        new JObject
                        {
                            { "toolName", definition.Name },
                            { "argumentName", property.Name },
                            { "expectedType", expectedType ?? string.Empty },
                            { "actualType", value.Type.ToString() }
                        });
                }

                if (string.Equals(expectedType, "array", StringComparison.OrdinalIgnoreCase))
                    ValidateArrayItems(definition, property.Name, value as JArray, propertySchema["items"] as JObject);
            }
        }

        private static void ValidateArrayItems(RevitMcpToolDefinition definition, string propertyName, JArray array, JObject itemsSchema)
        {
            if (array == null || itemsSchema == null)
                return;

            string expectedItemType = itemsSchema["type"] == null ? null : itemsSchema["type"].ToString();
            if (string.IsNullOrWhiteSpace(expectedItemType))
                return;

            for (int i = 0; i < array.Count; i++)
            {
                JToken item = array[i];
                if (!IsJsonTypeCompatible(item, expectedItemType))
                {
                    throw new RevitMcpToolArgumentException(
                        "Invalid MCP tool array item type for '" + propertyName + "' at index " + i + ". Expected " + expectedItemType + ", got " + (item == null ? "null" : item.Type.ToString()) + ".",
                        new JObject
                        {
                            { "toolName", definition.Name },
                            { "argumentName", propertyName },
                            { "itemIndex", i },
                            { "expectedType", expectedItemType },
                            { "actualType", item == null ? "null" : item.Type.ToString() }
                        });
                }
            }
        }

        private static bool IsJsonTypeCompatible(JToken value, string expectedType)
        {
            if (string.IsNullOrWhiteSpace(expectedType) || value == null)
                return true;

            switch (expectedType)
            {
                case "integer":
                    return value.Type == JTokenType.Integer;
                case "number":
                    return value.Type == JTokenType.Integer || value.Type == JTokenType.Float;
                case "string":
                    return value.Type == JTokenType.String;
                case "boolean":
                    return value.Type == JTokenType.Boolean;
                case "array":
                    return value.Type == JTokenType.Array;
                case "object":
                    return value.Type == JTokenType.Object;
                default:
                    return true;
            }
        }

        private static RevitMcpToolArgumentException CreateInvalidArgumentException(
            string propertyName,
            string expectedType,
            JToken actualValue,
            Exception innerException)
        {
            return new RevitMcpToolArgumentException(
                "Invalid MCP tool argument '" + propertyName + "'. Expected " + expectedType + ".",
                new JObject
                {
                    { "argumentName", propertyName ?? string.Empty },
                    { "expectedType", expectedType ?? string.Empty },
                    { "actualType", actualValue == null ? "null" : actualValue.Type.ToString() },
                    { "conversionError", innerException == null ? string.Empty : innerException.Message }
                });
        }

        private static JToken ToJToken(object value)
        {
            if (value == null)
                return JValue.CreateNull();

            if (value is JToken token)
                return token;

            string json = JsonConvert.SerializeObject(value);
            using (System.IO.StringReader stringReader = new System.IO.StringReader(json))
            using (JsonTextReader jsonReader = new JsonTextReader(stringReader))
            {
                jsonReader.DateParseHandling = DateParseHandling.None;
                return JToken.ReadFrom(jsonReader);
            }
        }

        private static RevitMcpToolCallResponse TryCreateToolResultErrorResponse(string toolName, JToken result)
        {
            JObject resultObject = result as JObject;
            if (resultObject == null)
                return null;

            JToken successToken = resultObject["success"];
            bool hasExplicitFailure = successToken != null
                && successToken.Type == JTokenType.Boolean
                && successToken.Value<bool>() == false;

            JToken errorToken = resultObject["error"];
            bool hasTopLevelError = errorToken != null
                && errorToken.Type != JTokenType.Null
                && !string.IsNullOrWhiteSpace(errorToken.Type == JTokenType.String
                    ? errorToken.Value<string>()
                    : JsonConvert.SerializeObject(errorToken));

            if (!hasExplicitFailure && !hasTopLevelError)
                return null;

            string message = hasTopLevelError
                ? (errorToken.Type == JTokenType.String ? errorToken.Value<string>() : JsonConvert.SerializeObject(errorToken))
                : "MCP tool returned an unsuccessful result.";

            return CreateErrorResponse(
                toolName,
                "tool_result_error",
                message,
                null,
                new JObject
                {
                    { "toolName", toolName ?? string.Empty },
                    { "result", resultObject.DeepClone() }
                });
        }

        private static RevitMcpToolCallResponse CreateErrorResponse(
            string toolName,
            string code,
            string message,
            string exceptionType,
            JObject details = null)
        {
            return new RevitMcpToolCallResponse
            {
                Success = false,
                ToolName = toolName,
                Error = new RevitMcpToolExecutionError
                {
                    Code = code,
                    Message = message,
                    ExceptionType = exceptionType,
                    Details = details
                }
            };
        }
    }
}


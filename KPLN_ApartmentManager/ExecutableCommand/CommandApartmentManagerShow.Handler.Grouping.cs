using Autodesk.Revit.DB;
using KPLN_ApartmentManager.Common;
using KPLN_ApartmentManager.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private void ApplyGeneratedElementsGrouping(Document doc, ViewPlan targetPlan, ApartmentGeneratedElementsGroupingMode groupingMode,
            Dictionary<long, ApartmentProcessState> apartmentStates, List<string> debugMessages)
        {
            if (doc == null || groupingMode == ApartmentGeneratedElementsGroupingMode.None ||
                apartmentStates == null || apartmentStates.Count == 0)
            {
                return;
            }

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Группировка построенных элементов"))
            {
                t.Start();
                ApplyApartmentFailureHandling(t);

                if (groupingMode == ApartmentGeneratedElementsGroupingMode.ByApartment)
                {
                    foreach (ApartmentProcessState state in apartmentStates.Values
                        .Where(x => x != null)
                        .OrderBy(x => IDHelper.ElIdValue(x.ApartmentId)))
                    {
                        CreateGeneratedElementsGroupForApartment(doc, state, debugMessages);
                    }
                }
                else if (groupingMode == ApartmentGeneratedElementsGroupingMode.WholePlan)
                {
                    CreateGeneratedElementsGroupForPlan(doc, targetPlan, apartmentStates, debugMessages);
                }

                t.Commit();
            }
        }

        private void CreateGeneratedElementsGroupForApartment(Document doc, ApartmentProcessState state, List<string> debugMessages)
        {
            if (state == null)
                return;

            List<ElementId> elementIds = CollectExistingCreatedElementIds(doc, state.CreatedElementIds);
            if (elementIds.Count < 2)
                return;

            string diagnostic;
            ElementId groupId;
            if (TryCreateGeneratedElementsGroup(
                doc,
                elementIds,
                "KPLN_AM_Квартира_" + IDHelper.ElIdValue(state.ApartmentId),
                out groupId,
                out diagnostic))
            {
                AddNavigationElementCandidate(state, groupId);
                return;
            }

            AddApartmentDiagnostic(state, debugMessages, diagnostic);
        }

        private void CreateGeneratedElementsGroupForPlan(Document doc, ViewPlan targetPlan, Dictionary<long, ApartmentProcessState> apartmentStates,
            List<string> debugMessages)
        {
            List<ElementId> elementIds = CollectExistingCreatedElementIds(
                doc,
                apartmentStates.Values
                    .Where(x => x != null && x.CreatedElementIds != null)
                    .SelectMany(x => x.CreatedElementIds));

            if (elementIds.Count < 2)
                return;

            string planName = targetPlan != null && !string.IsNullOrWhiteSpace(targetPlan.Name)
                ? targetPlan.Name
                : "План";

            string diagnostic;
            ElementId groupId;
            if (TryCreateGeneratedElementsGroup(
                doc,
                elementIds,
                "KPLN_AM_" + planName + "_Построено",
                out groupId,
                out diagnostic))
            {
                foreach (ApartmentProcessState state in apartmentStates.Values.Where(x => x != null))
                    AddNavigationElementCandidate(state, groupId);

                return;
            }

            ApartmentProcessState firstState = apartmentStates.Values.FirstOrDefault(x => x != null);
            AddApartmentDiagnostic(firstState, debugMessages, diagnostic);
        }

        private static List<ElementId> CollectExistingCreatedElementIds(Document doc, IEnumerable<ElementId> source)
        {
            List<ElementId> result = new List<ElementId>();
            HashSet<long> seen = new HashSet<long>();

            if (doc == null || source == null)
                return result;

            foreach (ElementId id in source)
            {
                if (id == null || id == ElementId.InvalidElementId)
                    continue;

                long value = IDHelper.ElIdValue(id);
                if (seen.Contains(value))
                    continue;

                Element element = doc.GetElement(id);
                if (element == null || element is ElementType)
                    continue;

                result.Add(id);
                seen.Add(value);
            }

            return result;
        }

        private static bool TryCreateGeneratedElementsGroup(Document doc, List<ElementId> elementIds, string baseGroupName,
            out ElementId groupId, out string diagnostic)
        {
            groupId = ElementId.InvalidElementId;
            diagnostic = null;

            if (doc == null || elementIds == null || elementIds.Count < 2)
            {
                diagnostic = "Группировка не выполнена: недостаточно построенных элементов.";
                return false;
            }

            try
            {
                Autodesk.Revit.DB.Group group = doc.Create.NewGroup(elementIds);
                if (group == null)
                {
                    diagnostic = "Revit не создал группу построенных элементов.";
                    return false;
                }

                groupId = group.Id;
                TryRenameGroupType(doc, group, baseGroupName);
                return true;
            }
            catch (Exception ex)
            {
                diagnostic = "Не удалось сгруппировать построенные элементы: " + ex.Message;
                return false;
            }
        }

        private static void TryRenameGroupType(Document doc, Autodesk.Revit.DB.Group group, string baseGroupName)
        {
            try
            {
                if (doc == null || group == null || group.GroupType == null)
                    return;

                Parameter nameParameter = group.GroupType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                if (nameParameter != null && !nameParameter.IsReadOnly)
                    nameParameter.Set(GenerateUniqueGroupTypeName(doc, baseGroupName));
            }
            catch
            {
            }
        }

        private static string GenerateUniqueGroupTypeName(Document doc, string baseGroupName)
        {
            string cleanBaseName = NormalizeGroupTypeName(baseGroupName);
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (GroupType groupType in new FilteredElementCollector(doc).OfClass(typeof(GroupType)).Cast<GroupType>())
                {
                    if (groupType != null && !string.IsNullOrWhiteSpace(groupType.Name))
                        usedNames.Add(groupType.Name);
                }
            }
            catch
            {
            }

            if (!usedNames.Contains(cleanBaseName))
                return cleanBaseName;

            for (int i = 2; i < 10000; i++)
            {
                string candidate = cleanBaseName + " (" + i + ")";
                if (!usedNames.Contains(candidate))
                    return candidate;
            }

            return cleanBaseName + " " + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        private static string NormalizeGroupTypeName(string value)
        {
            string result = string.IsNullOrWhiteSpace(value)
                ? "KPLN_AM_Построенные элементы"
                : value.Trim();

            result = result
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ");

            while (result.Contains("  "))
                result = result.Replace("  ", " ");

            return result.Length <= 120
                ? result
                : result.Substring(0, 120).Trim();
        }
    }
}

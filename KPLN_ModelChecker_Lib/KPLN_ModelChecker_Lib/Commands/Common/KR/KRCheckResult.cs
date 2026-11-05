using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_ModelChecker_Lib.Commands.Common.KR
{
    /// <summary>Full report for UI/MCP, plus standard CheckerEntity projections.</summary>
    public sealed class KRCheckResult
    {
        public JObject Data { get; }
        public string Status => (string)Data["status"];
        public KRCheckResult(IDictionary<string, object> data) { Data = JObject.FromObject(data); }
        public string ToJson() => Data.ToString(Formatting.None);

        internal IEnumerable<CheckerEntity> ToCheckerEntities(Document doc)
        {
            foreach (string section in new[] { "issues", "skipped_elements", "run_warnings" })
            foreach (var row in (Data[section] as JArray ?? new JArray()).OfType<JObject>())
            {
                bool isError = section == "issues" && (string)row["severity"] == "error";
                string header = section == "skipped_elements" ? "Не проверено"
                    : section == "run_warnings" ? "Ограничение" : isError ? "Ошибка" : "Требует разбора";
                string description = (string)(row["description"] ?? row["message"] ?? row["text"] ?? row["status"]) ?? "";
                var info = (JObject)row.DeepClone();
                bool sourceIsType = (string)row["class_source_kind"] == "type";
                string typeId = sourceIsType ? (string)row["class_source_id"] : null;
                if (sourceIsType) info.Remove("class_source_id");
                var ids = (row["element_ids"] as JArray ?? new JArray()).Select(x => (string)x).ToList();
                if (ids.Count == 0) ids.Add((string)(row["rebar_id"] ?? row["element_id"]));
                var elements = ids.Where(id => id != typeId).Distinct().Select(id => FindInstance(doc, id))
                    .Where(e => e != null).ToArray();
                if (sourceIsType && info["element_ids"] is JArray sourceIds)
                    info["element_ids"] = new JArray(sourceIds.Where(id => (string)id != typeId));
                CheckerEntity entity;
                try
                {
                    entity = elements.Length == 0
                        ? new CheckerEntity(header, description, info.ToString(Formatting.None))
                        : new CheckerEntity(elements, header, description, info.ToString(Formatting.None));
                }
                catch
                {
                    // Metadata needed by CheckerEntity may be unavailable; retain the full finding.
                    entity = new CheckerEntity(header, description, info.ToString(Formatting.None));
                }
                yield return entity.Set_Status(isError ? ErrorStatus.Error : ErrorStatus.Warning);
            }
        }

        private static Element FindInstance(Document doc, string text)
        {
            if (!long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) || id <= 0) return null;
            try
            {
#if Debug2024 || Revit2024
                var element = doc.GetElement(new ElementId(id));
#else
                if (id > int.MaxValue) return null;
                var element = doc.GetElement(new ElementId((int)id));
#endif
                return element != null && !(element is ElementType) && !(element is Material)
                    && !(element is PropertySetElement) && !(element is AppearanceAssetElement) ? element : null;
            }
            catch { return null; }
        }
    }
}

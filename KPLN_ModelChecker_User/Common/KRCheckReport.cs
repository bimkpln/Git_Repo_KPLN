using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_ModelChecker_User.Checks.KR
{
    internal sealed class KRCheckReportRow
    {
        public string Kind { get; set; }
        public string Element { get; set; }
        public string Description { get; set; }
        public string Actual { get; set; }
        public string Expected { get; set; }
        public string Host { get; set; }
        public string Mode { get; set; }
        public string Details { get; set; }
        internal string SourceId;
        internal string[] ElementIds;
    }

    internal sealed class KRCheckReport
    {
        internal JObject Data;
        internal string Summary;
        internal string Notes;
        internal readonly List<KRCheckReportRow> Rows = new List<KRCheckReportRow>();
        internal string Json => Data.ToString(Formatting.Indented);

        internal static string Text(JToken token) => token == null || token.Type == JTokenType.Null
            ? "" : token.Type == JTokenType.String ? (string)token : token.ToString(Formatting.None);

        private static readonly Dictionary<string, string> Labels = new Dictionary<string, string>
        {
            {"status", "Результат"}, {"description", "Описание"}, {"severity", "Категория"},
            {"rebar_id", "ID стержня"}, {"rebar_name", "Имя стержня"}, {"element_id", "ID элемента"},
            {"element_ids", "Элементы замечания"}, {"message", "Описание"},
            {"actual_mark", "Марка арматуры"}, {"expected_mark", "Марка основы"},
            {"rebar_mark", "Марка на самом стержне"}, {"actual_class", "Класс арматуры"},
            {"expected_class", "Ожидаемый класс"}, {"mode", "Способ проверки"},
            {"parent_id", "ID родителя"}, {"ancestor_ids", "Цепочка родителей"},
            {"mark_source_id", "Источник марки (ID)"}, {"mark_source_kind", "Источник марки"},
            {"class_source_id", "Источник класса (ID)"}, {"class_source_kind", "Источник класса"},
            {"class_source_name", "Имя источника класса"},
            {"selected_host_id", "Выбранная основа (ID)"}, {"host_id", "Основа (ID)"},
            {"host_name", "Имя основы"}, {"mark", "Марка основы"}, {"classes", "Классы бетона"},
            {"candidates", "Рассмотренные основы"}, {"geometry_notes", "Замечания к определению основы"},
            {"materials", "Материалы основы"}, {"material_id", "ID материала"},
            {"material_name", "Материал"}, {"code", "Код материала"}, {"class", "Класс бетона"},
            {"material_code", "Код материала"}, {"concrete_class", "Класс бетона"},
            {"intersection_volume_ft3", "Объём пересечения, фут³"},
            {"selected_intersection_volume_ft3", "Объём выбранного пересечения, фут³"},
            {"host_selection", "Выбор основы"}, {"parameter_read_only", "Параметр только для чтения"}
        };

        private static readonly Dictionary<string, string> Values = new Dictionary<string, string>
        {
            {"mark_mismatch", "Марки не совпадают"}, {"rebar_mark_empty", "Марка арматуры не заполнена"},
            {"host_mark_empty", "Марка основы не заполнена"}, {"parameter_missing", "Параметр марки недоступен или неоднозначен"},
            {"geometry_uncertain", "Неопределённость геометрии / материалов"}, {"host_not_found", "Основа не найдена"},
            {"no_host", "Основа не найдена"}, {"multiple_hosts", "Несколько основ"},
            {"check_failed", "Не удалось проверить элемент"}, {"class_mismatch", "Классы бетона не совпадают"},
            {"rebar_class_empty", "Класс арматуры не заполнен"}, {"rebar_class_invalid", "Некорректный класс арматуры"},
            {"host_class_uncertain", "Класс основы не определён однозначно"},
            {"parameter_uncertain", "Не удалось однозначно прочитать параметр"},
            {"parameter_read_failed", "Ошибка чтения параметра"},
            {"single_class", "Единый класс проекта"}, {"host", "Назначенная основа"},
            {"geometry", "Пересечение геометрии"}, {"host_or_geometry", "Назначенная основа / геометрия"},
            {"rebar", "Стержень"}, {"parent", "Родитель"}, {"instance", "Экземпляр"}, {"type", "Тип"},
            {"unknown", "Не определён"}, {"error", "Ошибка"}, {"review", "Требует разбора"},
            {"largest_intersection_then_lowest_id", "Максимальный объём пересечения; при равенстве меньший ID"},
            {"GetHostId", "Назначенная основа"}
        };

        internal static string Translate(string value) => Values.TryGetValue(value ?? "", out var result) ? result : value;

        internal static KRCheckReport Parse(string json) => FromData(JObject.Parse(json));

        internal static KRCheckReport FromData(JObject reportData)
        {
            var r = new KRCheckReport { Data = (JObject)reportData.DeepClone() };
            var d = r.Data;
            string status = Text(d["status"]);
            string conclusion = status == "partial" ? "Проверено частично — учитывайте ограничения и замечания."
                : status == "empty" ? "Нет элементов для проверки после отбора и исключений."
                : status == "completed" && (int?)d["error_count"] > 0 ? "Проверка завершена. Найдены ошибки."
                : status == "completed" ? "Проверка завершена. Ошибок не найдено в проверенном объёме."
                : "Проверка не завершена: " + status;
            r.Summary = "Модель: " + Text(d["model_title"]) + "\n" + conclusion + "\n"
                + "Проверено: " + Text(d["checked_count"]) + "   Совпало: " + Text(d["matched_count"])
                + "   Ошибок: " + Text(d["error_count"]) + "   Требуют разбора: " + Text(d["review_count"])
                + "   Не проверено: " + Text(d["unchecked_count"])
                + "\nИсключено по типу подсчёта: " + Text(d["excluded_counting_type_count"]);
            if (d["mode"] != null) r.Summary += "   Режим: " + Translate(Text(d["mode"]));
            if (d["project_concrete_classes"] is JArray classes)
                r.Summary += "\nКлассы бетона в модели: " + (classes.Count == 0 ? "не найдены" : string.Join(", ", classes.Select(Text)));
            r.Notes = string.Join("\n", (d["notes"] as JArray ?? new JArray()).Select(Text));
            foreach (var section in new[] { "issues", "skipped_elements", "run_warnings" })
            foreach (var item in d[section] as JArray ?? new JArray())
            {
                var data = item as JObject;
                if (data == null) continue;
                var isType = Text(data["class_source_kind"]) == "type";
                string actual = Text(data["actual_mark"] ?? data["actual_class"]);
                string expected = Text(data["expected_mark"] ?? data["expected_class"]);
                var row = new KRCheckReportRow
                {
                    Kind = section == "skipped_elements" ? "Не проверено" : section == "run_warnings" ? "Ограничение"
                        : Text(data["severity"]) == "error" ? "Ошибка" : "Требует разбора",
                    Element = Text(data["rebar_id"] ?? data["element_id"]),
                    Description = Translate(Text(data["status"])),
                    Actual = actual.Length == 0 ? "—" : "«" + actual + "»",
                    Expected = expected.Length == 0 ? "—" : "«" + expected + "»",
                    Host = Text(data["selected_host_id"]),
                    SourceId = isType ? "" : Text(data["mark_source_id"] ?? data["class_source_id"]),
                    Mode = Translate(Text(data["mode"])),
                    ElementIds = (data["element_ids"] as JArray ?? new JArray()).Select(Text)
                        .Where(id => !isType || id != Text(data["class_source_id"])).ToArray()
                };
                if (string.IsNullOrEmpty(row.Description)) row.Description = Text(data["description"] ?? data["message"]);
                if (row.ElementIds.Length == 0 && row.Element.Length > 0) row.ElementIds = new[] { row.Element };
                var details = new StringBuilder();
                AppendDetails(details, data, "");
                row.Details = details.ToString();
                r.Rows.Add(row);
            }
            return r;
        }

        private static void AppendDetails(StringBuilder text, JObject data, string indent)
        {
            foreach (var p in data.Properties())
            {
                // Types are parameter sources, not inspected/selected model instances.
                if (p.Name == "text" || p.Name == "element_ids" ||
                    (p.Name == "class_source_id" && Text(data["class_source_kind"]) == "type")) continue;
                string label = Labels.TryGetValue(p.Name, out var name) ? name : p.Name;
                if (p.Value is JArray array)
                {
                    text.AppendLine(indent + label + ":");
                    foreach (var child in array)
                        if (child is JObject obj) { AppendDetails(text, obj, indent + "  "); text.AppendLine(); }
                        else text.AppendLine(indent + "  " + Text(child));
                }
                else if (p.Value is JObject obj) { text.AppendLine(indent + label + ":"); AppendDetails(text, obj, indent + "  "); }
                else
                {
                    string value = Text(p.Value);
                    if (p.Name == "status" || p.Name == "mode" || p.Name.EndsWith("_kind") || p.Name == "host_selection" || p.Name == "severity")
                        value = Translate(value);
                    text.AppendLine(indent + label + ": " + (value.Length == 0 ? "—" : value));
                }
            }
        }
    }
}

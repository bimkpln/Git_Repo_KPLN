using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KPLN_RevitMcpBridge.Services
{
    // Adapter only: all selection, geometry and comparison rules remain in ModelChecker_User.
    // Called exclusively from the bridge's Revit API dispatcher.
    internal sealed class KRCheckService
    {
        private static readonly Dictionary<string, string> Types = new Dictionary<string, string>
        {
            { "CheckIfcRebarMarks", "KPLN_ModelChecker_User.Checks.KR.Ifc.Checker" },
            { "CheckRebarConcreteClass", "KPLN_ModelChecker_User.Checks.KR.Concrete.Entry" }
        };
        private static readonly string[] Sections = { "issues", "skipped_elements", "run_warnings" };
        private const int MaxReportChars = 64 * 1024 * 1024;
        private sealed class Saved
        {
            internal string Id, DocumentId, ModuleId;
            internal long Revision;
            internal int Size;
            internal DateTime Created;
            internal Dictionary<string, object> Report;
        }
        private readonly List<Saved> _results = new List<Saved>();

        private static Assembly Library()
        {
            var loaded = AppDomain.CurrentDomain.GetAssemblies()
                .Where(a => a.GetName().Name == "KPLN_ModelChecker_User").ToArray();
            if (loaded.Length != 1)
                throw new BridgeException("kr_library_unavailable",
                    "Нужна ровно одна загруженная KPLN_ModelChecker_User с проверками КР.", 409);
            return loaded[0];
        }

        private static MethodInfo CheckMethod(Assembly library, string checkId)
        {
            string typeName;
            if (!Types.TryGetValue(checkId, out typeName))
                throw new BridgeException("unknown_kr_check", "Неизвестная проверка КР.");
            var type = library.GetType(typeName);
            var method = type?.GetMethod("RunJson", BindingFlags.Public | BindingFlags.Static,
                null, new[] { typeof(UIApplication) }, null);
            if (method == null || method.ReturnType != typeof(string) ||
                !Equals(type.GetField("CheckId")?.GetRawConstantValue(), checkId))
                throw new BridgeException("kr_update_required",
                    "Загруженная KPLN_ModelChecker_User не содержит совместимую " + checkId + ". Обновите модуль и перезапустите Revit.", 409);
            return method;
        }

        private static void RequireAccess(Assembly library)
        {
            // Same gate as the ribbon commands; fail closed if the internal contract changes.
            var gate = library.GetType("KPLN_ModelChecker_User.Checks.KR.KRCheckAccess")?
                .GetMethod("IsAllowed", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic,
                    null, Type.EmptyTypes, null);
            if (gate == null || gate.ReturnType != typeof(bool))
                throw new BridgeException("kr_update_required", "Не найден совместимый контроль доступа КР.", 409);
            bool allowed;
            try { allowed = (bool)gate.Invoke(null, null); }
            catch { throw new BridgeException("kr_access_denied", "Не удалось подтвердить доступ пользователя КР/BIM.", 403); }
            if (!allowed)
                throw new BridgeException("kr_access_denied", "Проверки доступны только отделам КР и BIM.", 403);
        }

        internal object Catalog()
        {
            return Types.Keys.Select(id =>
            {
                var item = new Dictionary<string, object> {
                    { "check_id", id }, { "group", "Частные критерии проверки КР" },
                    { "title", id == "CheckIfcRebarMarks" ? "Марки IFC-арматуры" : "Класс бетона арматуры" },
                    { "available", false }
                };
                try
                {
                    var library = Library();
                    item["library_module_id"] = library.ManifestModule.ModuleVersionId.ToString("D");
                    item["library_version"] = library.GetName().Version.ToString();
                    CheckMethod(library, id);
                    RequireAccess(library);
                    item["available"] = true;
                }
                catch (BridgeException ex) { item["reason_code"] = ex.Code; item["reason"] = ex.Message; }
                return item;
            }).ToArray();
        }

        internal object Run(UIApplication app, IDictionary<string, object> input, string documentId, long revision)
        {
            if (Json.Integer(input.Get("expected_revision"), "expected_revision") != revision)
                throw new BridgeException("stale_document", "Модель изменилась; получите get_revit_context заново.", 409);
            var doc = app.ActiveUIDocument?.Document;
            if (doc == null || doc.IsFamilyDocument || doc.IsLinked)
                throw new BridgeException("invalid_document", "Откройте проект RVT.", 409);
            var id = input.Text("check_id", true);
            var library = Library();
            var method = CheckMethod(library, id);
            var moduleId = library.ManifestModule.ModuleVersionId.ToString("D");
            if (!string.Equals(input.Text("library_module_id", true), moduleId, StringComparison.OrdinalIgnoreCase))
                throw new BridgeException("kr_library_changed", "Сборка проверок изменилась; получите каталог заново.", 409);
            RequireAccess(library);
            string raw;
            try { raw = (string)method.Invoke(null, new object[] { app }); }
            catch (TargetInvocationException ex)
            { throw new BridgeException("kr_check_failed", (ex.InnerException ?? ex).Message); }
            if (raw == null || raw.Length > MaxReportChars)
                throw new BridgeException("kr_report_too_large", "Отчёт отсутствует или превышает лимит 64 млн символов; проверка не может считаться успешно прочитанной.");
            Dictionary<string, object> report;
            try { report = Json.Parse(raw, MaxReportChars); }
            catch (BridgeException) { throw new BridgeException("kr_invalid_report", "Проверка вернула некорректный JSON."); }
            if (!Equals(report.Get("check_id"), id) || !(report.Get("status") is string) ||
                Sections.Any(s => !(report.Get(s) is IList)))
                throw new BridgeException("kr_invalid_report", "Несовместимая структура отчёта КР.");
            var saved = new Saved {
                Id = Guid.NewGuid().ToString("N"), DocumentId = documentId, Revision = revision,
                ModuleId = moduleId, Created = DateTime.UtcNow, Report = report, Size = raw.Length
            };
            Prune();
            while (_results.Count >= 4 || _results.Sum(r => (long)r.Size) + saved.Size > MaxReportChars)
                _results.RemoveAt(0);
            _results.Add(saved);
            var summary = report.Where(p => !Sections.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value);
            summary["result_id"] = saved.Id;
            summary["document_id"] = documentId;
            summary["revision_at_run"] = revision;
            summary["library_module_id"] = moduleId;
            summary["expires_utc"] = saved.Created.AddMinutes(30).ToString("o");
            summary["section_counts"] = Sections.ToDictionary(s => s, s => ((IList)report[s]).Count);
            return summary;
        }

        internal object Page(IDictionary<string, object> input, string documentId, long currentRevision)
        {
            RequireAccess(Library());
            Prune();
            var resultId = input.Text("result_id", true);
            var saved = _results.SingleOrDefault(r => r.Id == resultId);
            if (saved == null)
                throw new BridgeException("kr_result_expired", "Результат не найден: хранится до 30 минут, максимум 4 отчёта / 64 млн символов; остановка моста очищает результаты.", 404);
            if (saved.DocumentId != documentId)
                throw new BridgeException("wrong_document", "Результат относится к другому документу.", 409);
            var section = input.Text("section", true);
            if (!Sections.Contains(section)) throw new BridgeException("invalid_input", "Неизвестный раздел отчёта.");
            var offset = input.Range("offset", 0, 0, int.MaxValue);
            var limit = input.Range("limit", 50, 1, 200);
            var rows = (IList)saved.Report[section];
            var items = rows.Cast<object>().Skip(offset).Take(limit).ToArray();
            return new {
                result_id = saved.Id, document_id = saved.DocumentId, check_id = saved.Report["check_id"],
                revision_at_run = saved.Revision, current_revision = currentRevision,
                is_stale = saved.Revision != currentRevision, library_module_id = saved.ModuleId,
                status = saved.Report["status"], section, offset, total = rows.Count, items,
                next_offset = offset + items.Length, has_more = (long)offset + items.Length < rows.Count
            };
        }

        private void Prune() { _results.RemoveAll(r => r.Created < DateTime.UtcNow.AddMinutes(-30)); }
    }
}

using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Тонкий адаптер к уже загруженной Пакетной выдаче; логика печати остаётся в Publication.</summary>
    internal static class PublicationPrintService
    {
        internal static object Print(Document doc, IDictionary<string, object> input)
        {
            var assembly = AppDomain.CurrentDomain.GetAssemblies().SingleOrDefault(a => a.GetName().Name == "KPLN_Publication");
            var service = assembly?.GetType("KPLN_Publication.BatchPrintService");
            var settingsType = assembly?.GetType("KPLN_Publication.YayPrintSettings");
            if (service == null || settingsType == null || !Equals(service.GetField("ContractVersion")?.GetRawConstantValue(), 1))
                throw new BridgeException("publication_update_required", "Загрузите KPLN_Publication с BatchPrintService v1 и перезапустите Revit.");

            if (input.Get("settings") != null && !(input.Get("settings") is IDictionary<string, object>))
                throw new BridgeException("invalid_settings", "Настройки печати должны быть объектом.");
            var values = input.Get("settings") as IDictionary<string, object> ?? new Dictionary<string, object>();
            var defaults = new Dictionary<string, object>
            {
                { "printerName", "PDFCreator" }, { "outputPDFFolder", @"C:\PDF_Print" },
                { "pdfNameConstructor", "<Номер листа>_<Имя листа>.pdf" },
                { "hiddenLineProcessing", "VectorProcessing" }, { "colorsType", "Color" }, { "rasterQuality", "Medium" },
                { "isPDFExport", true }, { "isPrintToPaper", false }, { "isMergePdfs", false },
                { "isUseOrientation", false }, { "isRefreshSchedules", true }, { "isExcludeBorders", false }, { "isDWGExport", false }
            };
            foreach (var pair in values)
            {
                if (!defaults.ContainsKey(pair.Key))
                    throw new BridgeException("invalid_settings", "Неизвестная настройка печати: " + pair.Key);
                defaults[pair.Key] = pair.Value;
            }

            var settings = Activator.CreateInstance(settingsType);
            foreach (var pair in defaults)
            {
                var field = settingsType.GetField(pair.Key);
                if (field == null)
                    throw new BridgeException("publication_update_required", "Контракт настроек Publication изменился: " + pair.Key);
                object value = pair.Value;
                if (field.FieldType.IsEnum && value is string)
                {
                    try { value = Enum.Parse(field.FieldType, (string)value, false); }
                    catch (ArgumentException) { throw new BridgeException("invalid_settings", "Неизвестный режим: " + pair.Key); }
                }
                if (value == null || !field.FieldType.IsInstanceOfType(value))
                    throw new BridgeException("invalid_settings", "Неверный тип настройки: " + pair.Key);
                field.SetValue(settings, value);
            }

            try
            {
                var result = (Dictionary<string, object>)service.GetMethod("PrintAllSheets").Invoke(null,
                    new[] { (object)doc, settings, input.Flag("dry_run", true), input.Range("timeout_seconds", 30, 5, 60) });
                result["document_id"] = input.Text("document_id", true);
                result["revision_at_request"] = Json.Integer(input.Get("expected_revision"), "expected_revision");
                result["publication_version"] = assembly.GetName().Version.ToString();
                result["document_title"] = doc.Title;
                if (result.ContainsKey("output_directory"))
                {
                    string manifest = Path.Combine((string)result["output_directory"], "manifest.json");
                    result["manifest_path"] = manifest;
                    File.WriteAllText(manifest, Json.Serialize(result), new UTF8Encoding(false));
                }
                return result;
            }
            catch (TargetInvocationException ex)
            {
                throw new BridgeException("publication_print_failed", (ex.InnerException ?? ex).Message);
            }
        }
    }
}

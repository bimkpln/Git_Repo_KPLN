using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_CoordiantorAI.Common;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitFailureMonitor : IDisposable
    {
        private readonly UIControlledApplication _application;
        private bool _disposed;
        private DateTime _lastLogUtc;
        internal RevitFailureStore Store { get; private set; }

        internal RevitFailureMonitor(UIControlledApplication application)
        {
            _application = application;
            Store = new RevitFailureStore(ModuleData.RevitVersion, RevitMcpDiagnosticLogger.LogException);
            try
            {
                application.ControlledApplication.FailuresProcessing += OnFailuresProcessing;
                application.ViewActivated += OnViewActivated;
                application.Idling += OnIdling;
                application.ApplicationClosing += OnApplicationClosing;
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        private static string Identify(Document document)
        {
            // Revit guarantees the same hash for wrappers of the same open document,
            // and a new hash when that file is opened again. The store adds the session id.
            return "doc-" + document.GetHashCode().ToString("X8", CultureInfo.InvariantCulture);
        }

        private void UpdateActive(Document document)
        {
            if (document == null || !document.IsValidObject)
                Store.SetActiveDocument(null, null);
            else
                Store.SetActiveDocument(Identify(document), Clip(document.Title, 256));
        }

        private void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {
            try { UpdateActive(e.CurrentActiveView == null ? null : e.CurrentActiveView.Document); }
            catch (Exception ex) { LogCaptureError(ex); }
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            try
            {
                UIApplication app = sender as UIApplication;
                if (app != null)
                    UpdateActive(app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document);
            }
            catch (Exception ex) { LogCaptureError(ex); }
        }

        private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            if (_disposed) return;
            try
            {
                // Copy plain values here; accessor/API objects are invalid after this callback.
                FailuresAccessor accessor = e.GetFailuresAccessor();
                Document doc = accessor.GetDocument();
                if (doc == null) return;
                string documentId = Identify(doc);
                string documentName = Clip(doc.Title, 256);
                string path = Clip(doc.PathName, 1024);
                string transactionName = Clip(accessor.GetTransactionName(), 512);
                DateTimeOffset now = DateTimeOffset.UtcNow;
                string eventId = Guid.NewGuid().ToString("N");
                var seen = new Dictionary<string, JObject>(StringComparer.Ordinal);
                IList<FailureMessageAccessor> messages = accessor.GetFailureMessages();
                Store.NoteDropped(Math.Max(0, messages.Count - 200));
                foreach (FailureMessageAccessor failure in messages.Take(200))
                {
                    try
                    {
                        JArray readErrors = new JArray();
                        string description = ReadValue(() => failure.GetDescriptionText(), "", readErrors, "description") ?? "";
                        IList<ElementId> failing = ReadValue<IList<ElementId>>(
                            () => failure.GetFailingElementIds().ToList(), new List<ElementId>(), readErrors, "element_ids");
                        IList<ElementId> additional = ReadValue<IList<ElementId>>(
                            () => failure.GetAdditionalElementIds().ToList(), new List<ElementId>(), readErrors, "additional_element_ids");
                        string failureId = ReadValue(() => failure.GetFailureDefinitionId().Guid.ToString(),
                            "", readErrors, "failure_definition_id");
                        string severity = ReadValue(() => failure.GetSeverity().ToString(), "Unknown", readErrors, "severity");
                        JObject record = new JObject
                        {
                            { "capture_event_id", eventId },
                            { "observed_at_utc", now.ToString("o", CultureInfo.InvariantCulture) },
                            { "time_moscow", now.ToOffset(TimeSpan.FromHours(3)).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
                            { "document_id", documentId },
                            { "document_name", documentName },
                            { "document_path", path },
                            { "transaction_name", transactionName },
                            { "failure_definition_id", failureId },
                            { "severity", severity },
                            { "description", Clip(description, 8192) },
                            { "description_truncated", description.Length > 8192 },
                            { "element_ids", new JArray(failing.Take(200).Select(IDHelper.ElIdValue)) },
                            { "additional_element_ids", new JArray(additional.Take(200).Select(IDHelper.ElIdValue)) },
                            { "element_ids_truncated", failing.Count > 200 || additional.Count > 200 },
                            { "occurrences_in_event", 1 },
                            { "capture_incomplete", readErrors.Count > 0 },
                            { "read_errors", readErrors }
                        };
                        string key = record["failure_definition_id"] + "|" + record["description"] + "|" +
                            string.Join(",", failing.Take(200).Select(IDHelper.ElIdValue).OrderBy(id => id)) + "|" +
                            string.Join(",", additional.Take(200).Select(IDHelper.ElIdValue).OrderBy(id => id));
                        if (failing.Count > 200 || additional.Count > 200 || description.Length > 8192 || readErrors.Count > 0)
                            key += Guid.NewGuid().ToString("N");
                        JObject duplicate;
                        if (seen.TryGetValue(key, out duplicate))
                            duplicate["occurrences_in_event"] = (int)duplicate["occurrences_in_event"] + 1;
                        else
                            seen.Add(key, record);
                    }
                    catch (Exception ex) { LogCaptureError(ex); }
                }
                foreach (JObject record in seen.Values)
                    Store.Add(record);
                // Do not resolve/delete failures or change the transaction processing result.
            }
            catch (Exception ex) { LogCaptureError(ex); }
        }

        private static string Clip(string value, int max)
        {
            value = value ?? string.Empty;
            return value.Length <= max ? value : value.Substring(0, max);
        }

        private static T ReadValue<T>(Func<T> read, T fallback, JArray errors, string field)
        {
            try { return read(); }
            catch (Exception ex)
            {
                errors.Add(field + ": " + Clip(ex.Message, 256));
                return fallback;
            }
        }

        private void LogCaptureError(Exception ex)
        {
            if (DateTime.UtcNow - _lastLogUtc < TimeSpan.FromMinutes(1)) return;
            _lastLogUtc = DateTime.UtcNow;
            RevitMcpDiagnosticLogger.LogException("Revit failure capture could not read some data", ex);
        }

        private void OnApplicationClosing(object sender, ApplicationClosingEventArgs e) { Dispose(); }
        private void OnProcessExit(object sender, EventArgs e) { Store.Dispose(); }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                _application.ControlledApplication.FailuresProcessing -= OnFailuresProcessing;
                _application.ViewActivated -= OnViewActivated;
                _application.Idling -= OnIdling;
                _application.ApplicationClosing -= OnApplicationClosing;
                AppDomain.CurrentDomain.ProcessExit -= OnProcessExit;
            }
            catch (Exception ex) { LogCaptureError(ex); }
            finally { Store.Dispose(); }
        }

    }
}

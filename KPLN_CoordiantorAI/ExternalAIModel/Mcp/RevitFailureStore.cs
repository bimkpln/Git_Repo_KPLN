using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    // Contains only detached data: safe to query while Revit's UI/API is blocked.
    internal sealed class RevitFailureStore : IDisposable
    {
        internal const string ToolName = "get_recent_revit_failures";
        internal const int PageBytes = 200 * 1024;
        private const int MaxBufferBytes = 4 * 1024 * 1024;
        private readonly object _sync = new object();
        private readonly object _fileSync = new object();
        private readonly List<JObject> _records = new List<JObject>();
        private readonly Action<string, Exception> _log;
        private readonly int _version;
        private readonly int _pid;
        private readonly string _sessionId = Guid.NewGuid().ToString("N");
        private long _sequence;
        private long _revision;
        private long _dropped;
        private int _bytes;
        private bool _disposed;
        private bool _writerRunning;
        private string _activeId;
        private string _activeName;
        private string _fileError;
        internal string FilePath { get; private set; }

        internal RevitFailureStore(int version, Action<string, Exception> log, string directory = null)
        {
            _version = version;
            _log = log;
            using (Process process = Process.GetCurrentProcess())
            {
                _pid = process.Id;
                directory = directory ?? Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KPLN", "CoordinatorAI", "RevitFailures");
                FilePath = Path.Combine(directory, string.Format(CultureInfo.InvariantCulture,
                    "failures_{0}_{1}_{2}_{3}.jsonl", version, _pid,
                    process.StartTime.ToUniversalTime().Ticks, _sessionId));
            }
            // No file/directory is created until an actual failure has been captured.
            CleanupAbandonedFiles(directory, log);
        }

        internal void SetActiveDocument(string id, string name)
        {
            lock (_sync)
            {
                _activeId = id;
                _activeName = name;
            }
        }

        internal void Add(JObject captured)
        {
            lock (_sync)
            {
                if (_disposed) return;
                JObject record = (JObject)captured.DeepClone();
                record["sequence"] = ++_sequence;
                record["session_id"] = _sessionId;
                record["process_id"] = _pid;
                record["revit_version"] = _version;
                int size = Size(record);
                if (size > PageBytes - 4096)
                {
                    _dropped++;
                    return;
                }
                _records.Add(record);
                _bytes += size;
                string docId = (string)record["document_id"];
                while (_records.Count(r => (string)r["document_id"] == docId) > 200)
                    RemoveAt(_records.FindIndex(r => (string)r["document_id"] == docId));
                while (_records.Count > 2000 || _bytes > MaxBufferBytes)
                    RemoveAt(0);
                _revision++;
                if (!_writerRunning)
                {
                    _writerRunning = true;
                    _ = Task.Run((Action)WriteSnapshots);
                }
            }
        }

        private void RemoveAt(int index)
        {
            _bytes -= Size(_records[index]);
            _records.RemoveAt(index);
            _dropped++;
        }

        internal void NoteDropped(int count)
        {
            lock (_sync) { if (!_disposed) _dropped += count; }
        }

        internal RevitMcpToolCallResponse Execute(JObject arguments, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                JObject a = arguments ?? new JObject();
                string[] allowed = { "document_id", "all_documents", "dateTime", "endDateTime",
                    "severity", "limit", "offset", "snapshot_sequence" };
                foreach (JProperty p in a.Properties())
                    if (!allowed.Contains(p.Name))
                        throw new ArgumentException("Unknown argument: " + p.Name);
                int limit = (int)Integer(a, "limit", 10, 1, 200);
                int offset = (int)Integer(a, "offset", 0, 0, int.MaxValue);
                long snapshot = Integer(a, "snapshot_sequence", long.MaxValue, 0, long.MaxValue);
                bool all = false;
                if (a["all_documents"] != null)
                {
                    if (a["all_documents"].Type != JTokenType.Boolean)
                        throw new ArgumentException("all_documents must be a boolean.");
                    all = (bool)a["all_documents"];
                }
                string docId = Text(a, "document_id");
                if (all && docId != null)
                    throw new ArgumentException("Use either document_id or all_documents, not both.");
                string severity = Text(a, "severity");
                if (severity != null && severity != "Warning" && severity != "Error" && severity != "DocumentCorruption")
                    throw new ArgumentException("severity must be Warning, Error or DocumentCorruption.");
                DateTimeOffset? start = ParseTime(Text(a, "dateTime"));
                DateTimeOffset? end = ParseTime(Text(a, "endDateTime"));
                if (start.HasValue && end.HasValue && start > end)
                    throw new ArgumentException("dateTime must not be later than endDateTime.");

                lock (_sync)
                {
                    if (_disposed) return Error("failure_capture_stopped", "Failure capture has stopped.");
                    if (!all && docId == null) docId = _activeId;
                    if (!all && docId == null)
                        return Error("active_document_unknown",
                            "No active document is known. Use all_documents=true to read captured errors from this process.");
                    if (!all && docId != _activeId && !_records.Any(r => (string)r["document_id"] == docId))
                        return Error("document_not_found", "No retained failures for this session-local document_id.");

                    snapshot = Math.Min(snapshot, _sequence);
                    List<JObject> filtered = _records.Where(r =>
                        (long)r["sequence"] <= snapshot &&
                        (all || (string)r["document_id"] == docId) &&
                        (severity == null || (string)r["severity"] == severity) &&
                        (!start.HasValue || RecordedTime(r) >= start.Value) &&
                        (!end.HasValue || RecordedTime(r) <= end.Value))
                        .OrderByDescending(r => (long)r["sequence"]).ToList();

                    JArray page = new JArray();
                    // Reserve space for pagination/context metadata.
                    int usedBytes = 4096;
                    foreach (JObject r in filtered.Skip(offset).Take(limit))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int size = Size(r) + 1;
                        if (usedBytes + size > PageBytes) break;
                        page.Add(r.DeepClone());
                        usedBytes += size;
                    }
                    bool more = (long)offset + page.Count < filtered.Count;
                    JObject result = new JObject
                    {
                        { "failures", page },
                        { "total_count", filtered.Count },
                        { "returned_count", page.Count },
                        { "limit", limit },
                        { "offset", offset },
                        { "has_more", more },
                        { "next_offset", more ? new JValue(offset + page.Count) : JValue.CreateNull() },
                        { "snapshot_sequence", snapshot },
                        { "session_id", _sessionId },
                        { "process_id", _pid },
                        { "revit_version", _version },
                        { "document_id", all ? null : docId },
                        { "active_document_id", _activeId },
                        { "dropped_count", _dropped },
                        { "buffer_incomplete", _dropped > 0 },
                        { "max_size_bytes", PageBytes },
                        { "file_write_error", _fileError },
                        { "note", "Captured observations only; not current failure status. Buffer limited to 200 records per document, 2000 total and 4 MiB. Element ids may be invalid after rollback." }
                    };
                    return new RevitMcpToolCallResponse
                    {
                        Success = true, ToolName = ToolName, Result = result,
                        RevitModelName = all ? null : docId == _activeId ? _activeName :
                            filtered.Select(r => (string)r["document_name"]).FirstOrDefault()
                    };
                }
            }
            catch (ArgumentException ex) { return Error("invalid_arguments", ex.Message); }
            catch (OverflowException) { return Error("invalid_arguments", "Integer argument is out of range."); }
        }

        internal static RevitMcpToolCallResponse Error(string code, string message)
        {
            return new RevitMcpToolCallResponse
            {
                Success = false, ToolName = ToolName,
                Error = new RevitMcpToolExecutionError { Code = code, Message = message }
            };
        }

        private static long Integer(JObject a, string key, long fallback, long min, long max)
        {
            JToken token = a[key];
            if (token == null) return fallback;
            if (token.Type != JTokenType.Integer)
                throw new ArgumentException(key + " must be an integer.");
            long value;
            if (!long.TryParse(token.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                throw new ArgumentException(key + " is out of range.");
            if (value < min || value > max)
                throw new ArgumentException(key + " is out of range.");
            return value;
        }

        private static string Text(JObject a, string key)
        {
            if (a[key] == null) return null;
            if (a[key].Type != JTokenType.String || string.IsNullOrWhiteSpace((string)a[key]))
                throw new ArgumentException(key + " must be a non-empty string.");
            string value = ((string)a[key]).Trim();
            if (value.Length > 128) throw new ArgumentException(key + " is too long.");
            return value;
        }

        private static DateTimeOffset? ParseTime(string value)
        {
            if (value == null) return null;
            DateTime local;
            string[] plain = { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss.FFFFFFF", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF" };
            if (DateTime.TryParseExact(value, plain, CultureInfo.InvariantCulture, DateTimeStyles.None, out local))
                return new DateTimeOffset(DateTime.SpecifyKind(local, DateTimeKind.Unspecified), TimeSpan.FromHours(3));
            DateTimeOffset zoned;
            string[] iso = { "yyyy-MM-dd'T'HH:mm:sszzz", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz",
                "yyyy-MM-dd'T'HH:mm:ss'Z'", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'" };
            if (DateTimeOffset.TryParseExact(value, iso, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out zoned))
                return zoned;
            throw new ArgumentException("Invalid date/time. Use yyyy-MM-dd HH:mm:ss or ISO 8601.");
        }

        private static DateTimeOffset RecordedTime(JObject record)
        {
            return DateTimeOffset.Parse((string)record["observed_at_utc"], CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal);
        }

        private static int Size(JObject record)
        {
            // Revit 2023/2024 can load Json.NET 13.0.1, which lacks ToString(Formatting).
            return Encoding.UTF8.GetByteCount(JsonConvert.SerializeObject(record, Formatting.None));
        }

        private void WriteSnapshots()
        {
            try
            {
                while (true)
                {
                    lock (_fileSync)
                    {
                        List<JObject> snapshot;
                        long revision;
                        lock (_sync)
                        {
                            if (_disposed) { _writerRunning = false; return; }
                            snapshot = _records.Select(r => (JObject)r.DeepClone()).ToList();
                            revision = _revision;
                        }
                        Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                        using (var stream = new FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.Read))
                        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                            foreach (JObject record in snapshot)
                                writer.WriteLine(JsonConvert.SerializeObject(record, Formatting.None));
                        lock (_sync)
                        {
                            _fileError = null;
                            if (revision == _revision || _disposed) { _writerRunning = false; return; }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lock (_sync)
                {
                    _fileError = "Local snapshot could not be written; captured records remain available in memory.";
                    _writerRunning = false;
                }
                Report(_log, "Revit failure snapshot write failed", ex);
            }
        }

        public void Dispose()
        {
            lock (_sync)
            {
                if (_disposed) return;
                _disposed = true;
                _records.Clear();
                _bytes = 0;
                _activeId = null;
                _activeName = null;
            }
            // Serialize deletion with the writer so it cannot recreate the file after shutdown.
            lock (_fileSync)
            {
                try { if (File.Exists(FilePath)) File.Delete(FilePath); }
                catch (Exception ex) { Report(_log, "Revit failure snapshot deletion failed", ex); }
            }
        }

        internal static void CleanupAbandonedFiles(string directory, Action<string, Exception> log)
        {
            try
            {
                if (!Directory.Exists(directory)) return;
                foreach (string file in Directory.EnumerateFiles(directory, "failures_*.jsonl").Take(1000))
                {
                    string[] parts = Path.GetFileNameWithoutExtension(file).Split('_');
                    int version, pid;
                    long ticks;
                    Guid session;
                    if (parts.Length != 5 || parts[0] != "failures" ||
                        !int.TryParse(parts[1], out version) || !int.TryParse(parts[2], out pid) || pid <= 0 ||
                        !long.TryParse(parts[3], out ticks) || ticks <= 0 ||
                        !Guid.TryParseExact(parts[4], "N", out session)) continue;
                    bool exited;
                    try
                    {
                        using (Process process = Process.GetProcessById(pid))
                            exited = process.HasExited || process.StartTime.ToUniversalTime().Ticks != ticks;
                    }
                    catch (ArgumentException) { exited = true; }
                    catch { continue; } // Unknown ownership: preserve the file.
                    if (!exited) continue;
                    try { File.Delete(file); }
                    catch (Exception ex) { Report(log, "Abandoned Revit failure file deletion failed", ex); }
                }
            }
            catch (Exception ex) { Report(log, "Revit failure file cleanup failed", ex); }
        }

        private static void Report(Action<string, Exception> log, string message, Exception error)
        {
            try { if (log != null) log(message, error); } catch { }
        }
    }
}

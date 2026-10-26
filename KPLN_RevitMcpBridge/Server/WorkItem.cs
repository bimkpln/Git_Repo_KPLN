using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Server
{
    // State transitions are atomic: a queued timeout can never execute later.
    internal sealed class WorkItem
    {
        private int _state; // 0 queued, 1 running, 2 completed, 3 cancelled
        public readonly string Id;
        public readonly string Payload;
        public readonly bool IsReadOnly;
        public readonly DateTime Created = DateTime.UtcNow;
        public readonly TaskCompletionSource<object> Completion = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
        public WorkItem(string id, string payload, bool isReadOnly = false) { Id = id; Payload = payload; IsReadOnly = isReadOnly; }

        // Только явный список. Новые/неизвестные команды по умолчанию защищены как запись.
        internal static bool IsReadOnlyCommand(string command)
        {
            return ReadCommands.Contains(command ?? "");
        }

        private static readonly HashSet<string> ReadCommands = new HashSet<string>(StringComparer.Ordinal)
        {
            "get_context", "get_documents", "get_selection", "get_categories", "find_elements", "get_elements",
            "get_group_members", "get_sheet_contents", "get_documentation_catalog", "get_sheet_layout",
            "check_sheet_layout", "get_project_levels", "get_view_annotations", "get_plan_range",
            "get_mep_connections", "get_space_annotation_context", "get_dimension_references",
            "get_view_elements", "get_view_visibility", "get_schedule_data", "check_intersections",
            "get_views", "get_families", "get_family_types", "get_family_document", "get_links"
        };

        // Вызывать под _gate. Завершённое чтение можно повторить без побочных эффектов.
        // Записи, экспорт и незавершённые запросы не вытесняются раньше срока хранения.
        internal static void TrimCompletedReads(Dictionary<string, WorkItem> operations, int capacity)
        {
            int remove = operations.Count - capacity + 1;
            if (remove <= 0) return;

            foreach (var item in operations.Values.Where(i => i.IsReadOnly && i.Completion.Task.IsCompleted)
                .OrderBy(i => i.Created).Take(remove).ToArray())
                operations.Remove(item.Id);
        }
        public string State => new[] { "queued", "running", "completed", "cancelled" }[Volatile.Read(ref _state)];
        public bool Cancel()
        {
            if (Interlocked.CompareExchange(ref _state, 3, 0) != 0) return false;
            Completion.TrySetResult(Json.Error(new BridgeException("cancelled", "Запрос отменён до начала выполнения.", 408)));
            return true;
        }
        public void Execute(Func<object> action)
        {
            if (Interlocked.CompareExchange(ref _state, 1, 0) != 0) return;
            object result;
            try { result = action(); }
            catch (Exception ex) { result = Json.Error(ex); }
            Completion.TrySetResult(result);
            Interlocked.Exchange(ref _state, 2);
        }
    }
}

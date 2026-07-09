using System;
using System.Collections.Generic;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class PluginUsageTracker
    {
        private static readonly TimeSpan PendingTransactionWindow = TimeSpan.FromSeconds(60);
        private static readonly TimeSpan RecentExecutionTransactionWindow = TimeSpan.FromSeconds(60);

        private readonly object _lock = new object();
        private readonly Stack<PluginExecution> _executions = new Stack<PluginExecution>();
        private readonly List<PendingPluginTransaction> _pendingTransactions =
            new List<PendingPluginTransaction>();
        private readonly PluginUsageRepository _repository;
        private readonly PluginUsageSyncService _syncService;
        private readonly ErrorGuard _errorGuard;
        private readonly string _defaultTabName;
        private RecentPluginExecution _recentExecution;

        public PluginUsageTracker(
            PluginUsageRepository repository,
            PluginUsageSyncService syncService,
            ErrorGuard errorGuard,
            string defaultTabName)
        {
            _repository = repository;
            _syncService = syncService;
            _errorGuard = errorGuard;
            _defaultTabName = defaultTabName ?? string.Empty;
        }

        public IDisposable BeginExecution(string tabName, string buttonName)
        {
            return BeginExecution(tabName, string.Empty, buttonName);
        }

        public IDisposable BeginExecution(string tabName, string panelName, string buttonName)
        {
            return BeginExecution(tabName, panelName, buttonName, false);
        }

        internal IDisposable BeginExecution(
            string tabName,
            string panelName,
            string buttonName,
            bool includeRecentUnscopedTransactions)
        {
            PluginExecution execution = new PluginExecution(
                Guid.NewGuid().ToString("N"),
                string.IsNullOrWhiteSpace(tabName) ? _defaultTabName : tabName,
                panelName ?? string.Empty,
                buttonName ?? string.Empty);

            try
            {
                lock (_lock)
                {
                    _executions.Push(execution);
                }

                WriteEvent(execution, "PluginStarted", string.Empty, 0, 0, 0);
                if (includeRecentUnscopedTransactions)
                    FlushPendingTransactions(execution);

                return new PluginExecutionScope(this, execution);
            }
            catch (Exception exception)
            {
                try
                {
                    RemoveExecution(execution);
                }
                catch
                {
                }

                SafeQueueException("PluginUsage.BeginExecution", exception);
                return EmptyDisposable.Instance;
            }
        }

        public void RecordDocumentChanged(string[] transactionNames, int addedCount, int modifiedCount, int deletedCount)
        {
            try
            {
                PluginExecution execution = GetCurrentExecution();
                if (execution == null)
                    execution = GetRecentExecution();

                if (execution == null)
                {
                    CapturePendingTransaction(transactionNames, addedCount, modifiedCount, deletedCount);
                    return;
                }

                WriteTransactionEvents(execution, transactionNames, addedCount, modifiedCount, deletedCount);
            }
            catch (Exception exception)
            {
                SafeQueueException("PluginUsage.RecordDocumentChanged", exception);
            }
        }

        private void CapturePendingTransaction(
            string[] transactionNames,
            int addedCount,
            int modifiedCount,
            int deletedCount)
        {
            DateTime now = DateTime.Now;
            lock (_lock)
            {
                RemoveOldPendingTransactions(now);
                _pendingTransactions.Add(new PendingPluginTransaction(
                    now,
                    transactionNames,
                    addedCount,
                    modifiedCount,
                    deletedCount));
            }
        }

        private void FlushPendingTransactions(PluginExecution execution)
        {
            PendingPluginTransaction[] pendingTransactions;
            DateTime now = DateTime.Now;

            lock (_lock)
            {
                RemoveOldPendingTransactions(now);
                pendingTransactions = _pendingTransactions.ToArray();
                _pendingTransactions.Clear();
            }

            foreach (PendingPluginTransaction pendingTransaction in pendingTransactions)
            {
                WriteTransactionEvents(
                    execution,
                    pendingTransaction.TransactionNames,
                    pendingTransaction.AddedCount,
                    pendingTransaction.ModifiedCount,
                    pendingTransaction.DeletedCount);
            }
        }

        private void RemoveOldPendingTransactions(DateTime now)
        {
            _pendingTransactions.RemoveAll(item => now - item.EventTime > PendingTransactionWindow);
        }

        private void WriteTransactionEvents(
            PluginExecution execution,
            string[] transactionNames,
            int addedCount,
            int modifiedCount,
            int deletedCount)
        {
            if (transactionNames == null || transactionNames.Length == 0)
            {
                WriteEvent(execution, "Transaction", string.Empty, addedCount, modifiedCount, deletedCount);
                return;
            }

            foreach (string transactionName in transactionNames)
            {
                WriteEvent(execution, "Transaction", transactionName, addedCount, modifiedCount, deletedCount);
            }
        }

        private void EndExecution(PluginExecution execution)
        {
            RememberRecentExecution(execution);
            RemoveExecution(execution);
        }

        private void RemoveExecution(PluginExecution execution)
        {
            lock (_lock)
            {
                if (_executions.Count == 0)
                    return;

                PluginExecution current = _executions.Peek();
                if (ReferenceEquals(current, execution))
                {
                    _executions.Pop();
                    return;
                }

                PluginExecution[] items = _executions.ToArray();
                _executions.Clear();
                for (int i = items.Length - 1; i >= 0; i--)
                {
                    if (!ReferenceEquals(items[i], execution))
                        _executions.Push(items[i]);
                }
            }
        }

        private PluginExecution GetCurrentExecution()
        {
            lock (_lock)
            {
                return _executions.Count == 0 ? null : _executions.Peek();
            }
        }

        private PluginExecution GetRecentExecution()
        {
            DateTime now = DateTime.Now;
            lock (_lock)
            {
                if (_recentExecution == null)
                    return null;

                if (now - _recentExecution.CompletedAt > RecentExecutionTransactionWindow)
                {
                    _recentExecution = null;
                    return null;
                }

                return _recentExecution.Execution;
            }
        }

        private void RememberRecentExecution(PluginExecution execution)
        {
            if (execution == null)
                return;

            lock (_lock)
            {
                _recentExecution = new RecentPluginExecution(execution, DateTime.Now);
            }
        }

        private void WriteEvent(
            PluginExecution execution,
            string eventType,
            string transactionName,
            int addedCount,
            int modifiedCount,
            int deletedCount)
        {
            UserContextSnapshot userContext = UserContextSnapshot.Current();
            PluginUsageRecord record = PluginUsageRecord.Create(
                execution.RunId,
                eventType,
                execution.TabName,
                execution.PanelName,
                execution.ButtonName,
                transactionName,
                userContext,
                addedCount,
                modifiedCount,
                deletedCount);
            _repository.InsertEvent(record);
            _syncService.RequestSyncSoon();
        }

        private void SafeQueueException(string source, Exception exception)
        {
            try
            {
                _errorGuard?.QueueException(source, exception);
            }
            catch
            {
            }
        }

        private sealed class PluginExecution
        {
            public PluginExecution(string runId, string tabName, string panelName, string buttonName)
            {
                RunId = runId;
                TabName = tabName;
                PanelName = panelName;
                ButtonName = buttonName;
            }

            public string RunId { get; private set; }
            public string TabName { get; private set; }
            public string PanelName { get; private set; }
            public string ButtonName { get; private set; }
        }

        private sealed class PendingPluginTransaction
        {
            public PendingPluginTransaction(
                DateTime eventTime,
                string[] transactionNames,
                int addedCount,
                int modifiedCount,
                int deletedCount)
            {
                EventTime = eventTime;
                TransactionNames = transactionNames ?? new string[0];
                AddedCount = addedCount;
                ModifiedCount = modifiedCount;
                DeletedCount = deletedCount;
            }

            public DateTime EventTime { get; private set; }
            public string[] TransactionNames { get; private set; }
            public int AddedCount { get; private set; }
            public int ModifiedCount { get; private set; }
            public int DeletedCount { get; private set; }
        }

        private sealed class RecentPluginExecution
        {
            public RecentPluginExecution(PluginExecution execution, DateTime completedAt)
            {
                Execution = execution;
                CompletedAt = completedAt;
            }

            public PluginExecution Execution { get; private set; }
            public DateTime CompletedAt { get; private set; }
        }

        private sealed class PluginExecutionScope : IDisposable
        {
            private readonly PluginUsageTracker _tracker;
            private PluginExecution _execution;

            public PluginExecutionScope(PluginUsageTracker tracker, PluginExecution execution)
            {
                _tracker = tracker;
                _execution = execution;
            }

            public void Dispose()
            {
                try
                {
                    PluginExecution execution = _execution;
                    if (execution == null)
                        return;

                    _execution = null;
                    _tracker.EndExecution(execution);
                }
                catch
                {
                }
            }
        }

        private sealed class EmptyDisposable : IDisposable
        {
            public static readonly EmptyDisposable Instance = new EmptyDisposable();

            private EmptyDisposable()
            {
            }

            public void Dispose()
            {
            }
        }
    }
}
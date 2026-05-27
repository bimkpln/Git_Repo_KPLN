using System;
using System.Threading;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class CentralSyncService : IDisposable
    {
        private readonly UserDataRepository _repository;
        private readonly ErrorGuard _errorGuard;
        private Timer _timer;
        private int _isSyncing;

        public CentralSyncService(UserDataRepository repository, ErrorGuard errorGuard)
        {
            _repository = repository;
            _errorGuard = errorGuard;
        }

        public void Start()
        {
            if (_timer != null)
                return;

            TimeSpan startDelay = TimeSpan.FromSeconds(ModuleData.SyncStartDelaySeconds);
            TimeSpan interval = TimeSpan.FromSeconds(ModuleData.SyncIntervalSeconds);
            _timer = new Timer(OnTimer, null, startDelay, interval);
        }

        public void RequestSyncSoon()
        {
            Timer timer = _timer;
            if (timer == null)
                return;

            timer.Change(
                TimeSpan.FromSeconds(ModuleData.SyncAfterWriteDelaySeconds),
                TimeSpan.FromSeconds(ModuleData.SyncIntervalSeconds));
        }

        public void Dispose()
        {
            Timer timer = _timer;
            _timer = null;
            timer?.Dispose();
        }

        private void OnTimer(object state)
        {
            if (Interlocked.Exchange(ref _isSyncing, 1) == 1)
                return;

            try
            {
                _repository.TrySyncPendingToCentral(ModuleData.SyncBatchSize);
            }
            catch (Exception exception)
            {
                _errorGuard.QueueException("CentralSync", exception);
            }
            finally
            {
                Interlocked.Exchange(ref _isSyncing, 0);
            }
        }
    }
}
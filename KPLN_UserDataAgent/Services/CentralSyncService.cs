using System;
using System.Threading;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class CentralSyncService : IDisposable
    {
        private readonly UserDataRepository _repository;
        private readonly ErrorGuard _errorGuard;
        private readonly object _randomLock = new object();
        private readonly Random _random = new Random(Guid.NewGuid().GetHashCode());
        private Timer _timer;
        private int _isSyncing;
        private int _isSoonSyncScheduled;

        public CentralSyncService(UserDataRepository repository, ErrorGuard errorGuard)
        {
            _repository = repository;
            _errorGuard = errorGuard;
        }

        public void Start()
        {
            if (_timer != null)
                return;

            TimeSpan startDelay = GetDelayWithJitter(ModuleData.SyncStartDelaySeconds);
            TimeSpan interval = GetDelayWithJitter(ModuleData.SyncIntervalSeconds);
            _timer = new Timer(OnTimer, null, startDelay, interval);
        }

        public void RequestSyncSoon()
        {
            Timer timer = _timer;
            if (timer == null)
                return;

            if (Interlocked.Exchange(ref _isSoonSyncScheduled, 1) == 1)
                return;

            timer.Change(
                GetDelayWithJitter(ModuleData.SyncAfterWriteDelaySeconds),
                GetDelayWithJitter(ModuleData.SyncIntervalSeconds));
        }

        public void Dispose()
        {
            Timer timer = _timer;
            _timer = null;
            timer?.Dispose();
        }

        public void SyncNow(string source)
        {
            RunSync(source ?? "CentralSync.Manual");
        }

        private void OnTimer(object state)
        {
            RunSync("CentralSync.Timer");
        }

        private void RunSync(string source)
        {
            if (Interlocked.Exchange(ref _isSyncing, 1) == 1)
                return;

            try
            {
                _repository.TrySyncPendingToCentral(ModuleData.SyncBatchSize);
            }
            catch (Exception exception)
            {
                _errorGuard.QueueException(source, exception);
            }
            finally
            {
                Interlocked.Exchange(ref _isSoonSyncScheduled, 0);
                Interlocked.Exchange(ref _isSyncing, 0);
            }
        }

        private TimeSpan GetDelayWithJitter(int baseDelaySeconds)
        {
            int jitterSeconds = 0;
            if (ModuleData.SyncRandomJitterSeconds > 0)
            {
                lock (_randomLock)
                {
                    jitterSeconds = _random.Next(0, ModuleData.SyncRandomJitterSeconds + 1);
                }
            }

            return TimeSpan.FromSeconds(baseDelaySeconds + jitterSeconds);
        }
    }
}
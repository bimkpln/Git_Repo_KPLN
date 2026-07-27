using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_UserDataAgent.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_UserDataAgent
{
    public class Module : IExternalModule
    {
        private static UserDataRepository _repository;
        private static PluginUsageRepository _pluginRepository;
        private static ErrorGuard _errorGuard;
        private static CentralSyncService _syncService;
        private static PluginUsageSyncService _pluginSyncService;
        private static PluginUsageTracker _pluginUsageTracker;
        private static NonDefaultCommandBindingService _nonDefaultCommandBindingService;
        private static bool _isInitialized;

        public Result Close()
        {
            SafeRun("Module.Close.Sync", () => _syncService?.SyncNow("Закрытие Revit"));
            SafeRun("Module.Close.PluginSync", () => _pluginSyncService?.SyncNow("Закрытие Revit"));
            SafeDisposePluginUsage();
            SafeRun("Module.Close.DisposeSync", () => _syncService?.Dispose());

            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            try
            {
                ModuleData.RevitMainWindowHandle = application.MainWindowHandle;

                _repository = new UserDataRepository(ModuleData.LocalDatabasePath, ModuleData.CentralDatabasePath);
                _errorGuard = new ErrorGuard(ModuleData.ShowDebugErrors, _repository.InsertError);

                _errorGuard.Run("Module.Execute", () =>
                {
                    _repository.InitializeLocal();

                    if (_isInitialized)
                        return;

                    AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
                    TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

                    application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                    application.ControlledApplication.DocumentClosing += OnDocumentClosing;
                    application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                    application.ControlledApplication.DocumentSaving += OnDocumentSaving;
                    application.ControlledApplication.DocumentSaved += OnDocumentSaved;
                    application.ControlledApplication.DocumentSavingAs += OnDocumentSavingAs;
                    application.ControlledApplication.DocumentSavedAs += OnDocumentSavedAs;
                    application.ControlledApplication.DocumentSynchronizingWithCentral += OnDocumentSynchronizingWithCentral;
                    application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronizedWithCentral;
                    application.ViewActivated += OnViewActivated;
                    application.Idling += OnIdling;

                    _syncService = new CentralSyncService(_repository, _errorGuard);
                    _syncService.Start();

                    InitializePluginUsage(application, tabName);

                    _isInitialized = true;
                });
            }
            catch (Exception exception)
            {
                SafeQueueException("Module.Execute.Fatal", exception);
            }

            return Result.Succeeded;
        }

        public static IDisposable BeginPluginExecution(string tabName, string buttonName)
        {
            PluginUsageTracker tracker = _pluginUsageTracker;
            return tracker == null
                ? EmptyDisposable.Instance
                : tracker.BeginExecution(tabName, buttonName);
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs args)
        {
            Exception exception = args.ExceptionObject as Exception
                ?? new Exception(args.ExceptionObject?.ToString() ?? "Unknown unhandled exception");

            SafeQueueException("AppDomain.UnhandledException", exception);
        }

        private static void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs args)
        {
            SafeQueueException("TaskScheduler.UnobservedTaskException", args.Exception);
            args.SetObserved();
        }

        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            WriteEvent("DocumentOpened", args.Document);
        }

        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs args)
        {
            WriteEvent("DocumentClosing", args.Document);
        }

        private static void OnDocumentSaving(object sender, DocumentSavingEventArgs args)
        {
            WriteEvent("DocumentSaving", args.Document);
        }

        private static void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            WriteEvent("DocumentSaved", args.Document);
        }

        private static void OnDocumentSavingAs(object sender, DocumentSavingAsEventArgs args)
        {
            WriteEvent("DocumentSavingAs", args.Document);
        }

        private static void OnDocumentSavedAs(object sender, DocumentSavedAsEventArgs args)
        {
            WriteEvent("DocumentSavedAs", args.Document);
        }

        private static void OnDocumentSynchronizingWithCentral(object sender, DocumentSynchronizingWithCentralEventArgs args)
        {
            WriteEvent("DocumentSynchronizingWithCentral", args.Document);
        }

        private static void OnDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            WriteEvent("DocumentSynchronizedWithCentral", args.Document);
        }

        private static void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            WriteEvent("ViewActivated", args.Document);
        }

        private static void OnIdling(object sender, IdlingEventArgs args)
        {
            SafeRun("Module.PluginUsage.Idling", () => _nonDefaultCommandBindingService?.OnIdling());
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            Document document = null;
            string eventName = "DocumentChanged";
            string[] transactions = new string[0];
            int addedCount = 0;
            int modifiedCount = 0;
            int deletedCount = 0;

            SafeRun("DocumentChanged.Read", () =>
            {
                document = args.GetDocument();

                try
                {
                    eventName = string.Format("DocumentChanged.{0}", args.Operation);
                }
                catch
                {
                }

                IEnumerable<string> transactionNames = args.GetTransactionNames();
                transactions = transactionNames == null
                    ? new string[0]
                    : transactionNames.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct().ToArray();

                addedCount = SafeCount(() => args.GetAddedElementIds());
                modifiedCount = SafeCount(() => args.GetModifiedElementIds());
                deletedCount = SafeCount(() => args.GetDeletedElementIds());
            });

            SafeRun("DocumentChanged.UserEvents", () =>
            {
                if (document == null)
                    return;

                if (transactions.Length == 0)
                {
                    WriteEventUnsafe(eventName, document, string.Empty, addedCount, modifiedCount, deletedCount);
                }
                else
                {
                    foreach (string transactionName in transactions)
                    {
                        WriteEventUnsafe(eventName, document, transactionName, addedCount, modifiedCount, deletedCount);
                    }
                }
            });

            SafeRun("DocumentChanged.PluginUsage", () =>
            {
                _pluginUsageTracker?.RecordDocumentChanged(
                    transactions,
                    addedCount,
                    modifiedCount,
                    deletedCount);
            });
        }

        private static void WriteEvent(string eventName, Document document)
        {
            _errorGuard.Run(eventName, () => WriteEventUnsafe(eventName, document, string.Empty, 0, 0, 0));
        }

        private static void WriteEventUnsafe(
            string eventName,
            Document document,
            string transactionName,
            int addedCount,
            int modifiedCount,
            int deletedCount)
        {
            if (document == null)
                return;

            UserContextSnapshot userContext = UserContextSnapshot.Current();
            UserEventRecord record = UserEventRecord.Create(
                eventName,
                transactionName,
                userContext,
                addedCount,
                modifiedCount,
                deletedCount);
            _repository.InsertEvent(record);
            _syncService?.RequestSyncSoon();
        }

        private static int SafeCount(Func<ICollection<ElementId>> getIds)
        {
            try
            {
                ICollection<ElementId> ids = getIds();
                return ids == null ? 0 : ids.Count;
            }
            catch
            {
                return 0;
            }
        }

        private static void InitializePluginUsage(UIControlledApplication application, string tabName)
        {
            try
            {
                _pluginRepository = new PluginUsageRepository(
                    ModuleData.PluginLocalDatabasePath,
                    ModuleData.PluginCentralDatabasePath);
                _pluginRepository.InitializeLocal();
                _pluginSyncService = new PluginUsageSyncService(_pluginRepository, _errorGuard);
                _pluginUsageTracker = new PluginUsageTracker(
                    _pluginRepository,
                    _pluginSyncService,
                    _errorGuard,
                    tabName);
                _nonDefaultCommandBindingService = new NonDefaultCommandBindingService(
                    application,
                    _pluginUsageTracker,
                    _errorGuard);
                _pluginSyncService.Start();
            }
            catch (Exception exception)
            {
                SafeQueueException("Module.PluginUsage.Initialize", exception);
                SafeDisposePluginUsage();
            }
        }

        private static void SafeDisposePluginUsage()
        {
            SafeRun("Module.PluginUsage.DisposeCommands", () => _nonDefaultCommandBindingService?.Dispose());
            SafeRun("Module.PluginUsage.DisposeSync", () => _pluginSyncService?.Dispose());

            _nonDefaultCommandBindingService = null;
            _pluginUsageTracker = null;
            _pluginSyncService = null;
            _pluginRepository = null;
        }

        private static void SafeRun(string source, Action action)
        {
            try
            {
                action();
            }
            catch (Exception exception)
            {
                SafeQueueException(source, exception);
            }
        }

        private static void SafeQueueException(string source, Exception exception)
        {
            try
            {
                _errorGuard?.QueueException(source, exception);
            }
            catch
            {
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
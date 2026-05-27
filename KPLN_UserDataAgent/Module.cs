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
        private static ErrorGuard _errorGuard;
        private static CentralSyncService _syncService;
        private static bool _isInitialized;

        public Result Close()
        {
            _syncService?.Dispose();
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            _repository = new UserDataRepository(ModuleData.LocalDatabasePath, ModuleData.CentralDatabasePath);
            _errorGuard = new ErrorGuard(ModuleData.ShowDebugErrors);

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

                _isInitialized = true;
            });

            return Result.Succeeded;
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs args)
        {
            Exception exception = args.ExceptionObject as Exception
                ?? new Exception(args.ExceptionObject?.ToString() ?? "Unknown unhandled exception");

            _errorGuard?.HandleException("AppDomain.UnhandledException", exception);
        }

        private static void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs args)
        {
            _errorGuard?.HandleException("TaskScheduler.UnobservedTaskException", args.Exception);
            args.SetObserved();
        }

        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            WriteEvent("Документ открыт", args.Document);
        }

        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs args)
        {
            WriteEvent("Документ закрыт", args.Document);
        }

        private static void OnDocumentSaving(object sender, DocumentSavingEventArgs args)
        {
            WriteEvent("Сохранение документа", args.Document);
        }

        private static void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            WriteEvent("Документ сохранён", args.Document);
        }

        private static void OnDocumentSavingAs(object sender, DocumentSavingAsEventArgs args)
        {
            WriteEvent("Сохранить документ как", args.Document);
        }

        private static void OnDocumentSavedAs(object sender, DocumentSavedAsEventArgs args)
        {
            WriteEvent("Документ сохранён как", args.Document);
        }

        private static void OnDocumentSynchronizingWithCentral(object sender, DocumentSynchronizingWithCentralEventArgs args)
        {
            WriteEvent("Синхронизировать документ с ЦМ", args.Document);
        }

        private static void OnDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            WriteEvent("Документ синхронизирован с ЦМ", args.Document);
        }

        private static void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            WriteEvent("Активирован вид", args.Document);
        }

        private static void OnIdling(object sender, IdlingEventArgs args)
        {
            _errorGuard?.ShowPendingDialog();
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            _errorGuard.Run("DocumentChanged", () =>
            {
                Document document = args.GetDocument();
                if (document == null)
                    return;

                string eventName = "DocumentChanged";
                try
                {
                    eventName = string.Format("DocumentChanged.{0}", args.Operation);
                }
                catch
                {
                }

                IEnumerable<string> transactionNames = args.GetTransactionNames();
                string[] transactions = transactionNames == null
                    ? new string[0]
                    : transactionNames.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct().ToArray();

                int addedCount = SafeCount(() => args.GetAddedElementIds());
                int modifiedCount = SafeCount(() => args.GetModifiedElementIds());
                int deletedCount = SafeCount(() => args.GetDeletedElementIds());

                if (transactions.Length == 0)
                {
                    WriteEventUnsafe(eventName, document, string.Empty, addedCount, modifiedCount, deletedCount);
                    return;
                }

                foreach (string transactionName in transactions)
                {
                    WriteEventUnsafe(eventName, document, transactionName, addedCount, modifiedCount, deletedCount);
                }
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

            DocumentSnapshot documentSnapshot = DocumentSnapshot.FromDocument(document);
            UserContextSnapshot userContext = UserContextSnapshot.Current();
            UserEventRecord record = UserEventRecord.Create(
                eventName,
                transactionName,
                documentSnapshot,
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
    }
}
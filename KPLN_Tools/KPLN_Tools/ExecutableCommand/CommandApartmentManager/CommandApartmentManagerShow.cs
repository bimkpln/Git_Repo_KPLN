using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_Tools.ExecutableCommand
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandApartmentManagerShow : IExternalCommand
    {
        private static ApartmentManagerWindow _window;
        private static ApartmentManagerExternalController _controller;
        private static ApartmentPresetData _sessionPresetData;
        private static ApartmentPresetPanelContext _sessionPresetContext;
        private static Autodesk.Revit.ApplicationServices.Application _revitApplication;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;

                if (_window != null)
                {
                    ShowAndActivateExistingWindow();
                    return Result.Succeeded;
                }

                _controller = new ApartmentManagerExternalController();
                _window = new ApartmentManagerWindow(
                    SQLiteMainService.CurrentUserDBSubDepartment.Id,
                    _controller,
                    _sessionPresetData != null ? _sessionPresetData.Clone() : null,
                    _sessionPresetContext != null ? _sessionPresetContext.CloneSnapshot() : null);

                _controller.AttachWindow(_window);
                _window.Closed += OnWindowClosed;
                EnsureDocumentChangedSubscription(uiapp);

                new WindowInteropHelper(_window).Owner = uiapp.MainWindowHandle;

                _window.Show();
                _window.Activate();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private static void ShowAndActivateExistingWindow()
        {
            if (_window == null)
                return;

            if (!_window.IsVisible)
                _window.Show();

            if (_window.WindowState == WindowState.Minimized)
                _window.WindowState = WindowState.Normal;

            _window.Activate();
            _window.Topmost = true;
            _window.Topmost = false;
            _window.Focus();
        }

        private static void OnWindowClosed(object sender, EventArgs e)
        {
            if (_window != null)
            {
                _sessionPresetData = _window.ApartmentPresetData != null
                    ? _window.ApartmentPresetData.Clone()
                    : null;
                _sessionPresetContext = _window.GetApartmentPresetContextSnapshot();
            }

            if (_controller != null)
                _controller.DetachWindow();

            // Closed is raised outside Revit API context; keep the subscription and no-op while the window is null.
            _controller = null;
            _window = null;
        }

        private static void EnsureDocumentChangedSubscription(UIApplication uiapp)
        {
            if (uiapp == null || uiapp.Application == null)
                return;

            if (ReferenceEquals(_revitApplication, uiapp.Application))
                return;

            ClearDocumentChangedSubscription();

            _revitApplication = uiapp.Application;
            _revitApplication.DocumentChanged += RevitApplication_DocumentChanged;
        }

        private static void ClearDocumentChangedSubscription()
        {
            if (_revitApplication == null)
                return;

            _revitApplication.DocumentChanged -= RevitApplication_DocumentChanged;
            _revitApplication = null;
        }

        private static void RevitApplication_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            if (_window == null || e == null)
                return;

            ApartmentPresetData presetData = _window.ApartmentPresetData;
            HashSet<long> knownApartmentIds = ParseApartmentIdsFromSignature(
                presetData != null ? presetData.SelectedPlanModelSignature : null);

            if (knownApartmentIds.Count == 0)
                return;

            if (ContainsKnownApartmentId(e.GetDeletedElementIds(), knownApartmentIds) ||
                ContainsKnownApartmentId(e.GetModifiedElementIds(), knownApartmentIds) ||
                ContainsMarkedApartment(e.GetDocument(), e.GetAddedElementIds()))
            {
                _window.MarkApartmentPresetDataStale();
            }
        }

        private static HashSet<long> ParseApartmentIdsFromSignature(string signature)
        {
            HashSet<long> result = new HashSet<long>();
            if (string.IsNullOrWhiteSpace(signature))
                return result;

            foreach (string part in signature.Split(';'))
            {
                if (string.IsNullOrWhiteSpace(part))
                    continue;

                int markerIndex = part.IndexOf(":loggia=", StringComparison.OrdinalIgnoreCase);
                if (markerIndex <= 0)
                    continue;

                long id;
                if (long.TryParse(part.Substring(0, markerIndex), out id) && id > 0)
                    result.Add(id);
            }

            return result;
        }

        private static bool ContainsKnownApartmentId(ICollection<ElementId> ids, HashSet<long> knownApartmentIds)
        {
            if (ids == null || ids.Count == 0 || knownApartmentIds == null || knownApartmentIds.Count == 0)
                return false;

            foreach (ElementId id in ids)
            {
                if (id != null && knownApartmentIds.Contains(IDHelper.ElIdValue(id)))
                    return true;
            }

            return false;
        }

        private static bool ContainsMarkedApartment(Document doc, ICollection<ElementId> ids)
        {
            if (doc == null || ids == null || ids.Count == 0)
                return false;

            foreach (ElementId id in ids)
            {
                FamilyInstance familyInstance = doc.GetElement(id) as FamilyInstance;
                if (familyInstance == null)
                    continue;

                Parameter comments = familyInstance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                string value = comments != null ? comments.AsString() : null;
                if (!string.IsNullOrWhiteSpace(value) && value.Contains("[KPLN_APT_INSTANCE]"))
                    return true;
            }

            return false;
        }
    }

    internal class ApartmentManagerExternalController : IApartmentManagerExternalController
    {
        private readonly ApartmentManagerExternalHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public ApartmentManagerExternalController()
        {
            _handler = new ApartmentManagerExternalHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void AttachWindow(ApartmentManagerWindow window)
        {
            _handler.AttachWindow(window);
        }

        public void DetachWindow()
        {
            _handler.DetachWindow();
        }

        public void RequestPlaceApartment(int apartmentId)
        {
            _handler.PreparePlaceApartment(apartmentId);
            _externalEvent.Raise();
        }

        public void RequestConvertTo3D(ApartmentPresetData presetData)
        {
            _handler.PrepareConvertTo3D(presetData);
            _externalEvent.Raise();
        }

        public void RequestRefreshApartmentPresets(ApartmentPresetData presetData)
        {
            _handler.PrepareRefreshApartmentPresets(presetData);
            _externalEvent.Raise();
        }

        public void RequestUpdateApartmentMarks()
        {
            _handler.PrepareUpdateApartmentMarks();
            _externalEvent.Raise();
        }
    }
}

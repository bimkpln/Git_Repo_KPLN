using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Tools.Forms;
using System;
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

            _controller = null;
            _window = null;
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
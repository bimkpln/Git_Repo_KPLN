using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Forms.Commands;
using KPLN_Clashes_Ribbon.Forms.Entities;
using KPLN_Clashes_Ribbon.Services;
using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace KPLN_Clashes_Ribbon.Forms.ViewModels
{
    public sealed class ZoomSettingsVM
    {
        private readonly Window _mainWindow;
        private readonly Document _doc;
        private readonly DispatcherTimer _refreshTimer;
        private ZoomSettings _lastConfig;

        public ZoomSettingsVM(Window mainWindow, Document doc)
        {
            _mainWindow = mainWindow;
            _doc = doc;

            CurrentZoomSettingsM = new ZoomSettingsM();
            LoadConfigToWindow(true);

            SaveConfigCmd = new RelayCommand<object>(_ => SaveConfig(), _ => CurrentZoomSettingsM.CanSave);
            DropToDefaultCmd = new RelayCommand<object>(_ => DropToDefault());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);

            _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            _refreshTimer.Tick += (sender, args) => LoadConfigToWindow(false);
            _refreshTimer.Start();

            _mainWindow.Closed += (sender, args) => _refreshTimer.Stop();
        }

        public ZoomSettingsM CurrentZoomSettingsM { get; }

        public ICommand SaveConfigCmd { get; }

        public ICommand DropToDefaultCmd { get; }


        public ICommand CloseWindowCmd { get; }

        public void SaveConfig()
        {
            if (!CurrentZoomSettingsM.TryCreateConfig(out ZoomSettings settings, out string errorMessage))
            {
                MessageBox.Show(_mainWindow, errorMessage, "Настройки зума", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                ZoomSettingsConfigService.Save(settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show(_mainWindow, $"Не удалось сохранить shared-конфиг настроек зума:\n{ex.Message}", "Настройки зума", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _lastConfig = settings;
            CurrentZoomSettingsM.ApplyConfig(settings);
            _mainWindow.Close();
        }

        public void DropToDefault() => CurrentZoomSettingsM.DropToDefault();

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        private void LoadConfigToWindow(bool force)
        {
            ZoomSettings settings = ZoomSettingsConfigService.LoadOrCreateDefault(_doc);
            if (!force && CurrentZoomSettingsM.IsDirty)
                return;

            if (!force && _lastConfig != null && _lastConfig.Equals(settings))
                return;

            _lastConfig = settings;
            CurrentZoomSettingsM.ApplyConfig(settings);
        }
    }
}




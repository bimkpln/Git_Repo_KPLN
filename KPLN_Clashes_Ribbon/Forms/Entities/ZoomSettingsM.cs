using KPLN_Clashes_Ribbon.Core;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_Clashes_Ribbon.Forms.Entities
{
    public sealed class ZoomSettingsM : INotifyPropertyChanged
    {
        private bool _isApplyingConfig;
        private bool _isDirty;
        private string _offsetXMillimetersText;
        private string _offsetYMillimetersText;
        private string _offsetZMinMillimetersText;
        private string _offsetZMaxMillimetersText;
        private bool _fitCollision;
        private bool _canSave;
        private string _userHelp;

        public event PropertyChangedEventHandler PropertyChanged;

        public string OffsetXMillimetersText
        {
            get => _offsetXMillimetersText;
            set => SetTextValue(ref _offsetXMillimetersText, value);
        }

        public string OffsetYMillimetersText
        {
            get => _offsetYMillimetersText;
            set => SetTextValue(ref _offsetYMillimetersText, value);
        }

        public string OffsetZMinMillimetersText
        {
            get => _offsetZMinMillimetersText;
            set => SetTextValue(ref _offsetZMinMillimetersText, value);
        }

        public string OffsetZMaxMillimetersText
        {
            get => _offsetZMaxMillimetersText;
            set => SetTextValue(ref _offsetZMaxMillimetersText, value);
        }

        public bool FitCollision
        {
            get => _fitCollision;
            set
            {
                if (_fitCollision == value)
                    return;

                _fitCollision = value;
                MarkDirty();
                NotifyPropertyChanged();
            }
        }

        public bool CanSave
        {
            get => _canSave;
            private set
            {
                if (_canSave == value)
                    return;

                _canSave = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsDirty
        {
            get => _isDirty;
            private set
            {
                if (_isDirty == value)
                    return;

                _isDirty = value;
                NotifyPropertyChanged();
            }
        }

        public string UserHelp
        {
            get => _userHelp;
            private set
            {
                _userHelp = string.IsNullOrWhiteSpace(value) ? string.Empty : $"ВАЖНО: {value}";
                NotifyPropertyChanged();
            }
        }

        public void ApplyConfig(ZoomSettings settings, bool markDirty = false)
        {
            if (settings == null)
                settings = new ZoomSettings();

            _isApplyingConfig = true;
            OffsetXMillimetersText = FormatMillimeters(settings.OffsetXMillimeters);
            OffsetYMillimetersText = FormatMillimeters(settings.OffsetYMillimeters);
            OffsetZMinMillimetersText = FormatMillimeters(settings.OffsetZMinMillimeters);
            OffsetZMaxMillimetersText = FormatMillimeters(settings.OffsetZMaxMillimeters);
            FitCollision = settings.FitCollision;
            _isApplyingConfig = false;

            IsDirty = markDirty;
            UpdateCanSave();
        }

        public void DropToDefault() => ApplyConfig(new ZoomSettings(), true);

        public bool TryCreateConfig(out ZoomSettings settings, out string errorMessage)
        {
            settings = null;

            if (!TryGetMillimeters(OffsetXMillimetersText, "Расширение X", out double offsetX, out errorMessage)
                || !TryGetMillimeters(OffsetYMillimetersText, "Расширение Y", out double offsetY, out errorMessage)
                || !TryGetMillimeters(OffsetZMinMillimetersText, "Расширение Z вниз", out double offsetZMin, out errorMessage)
                || !TryGetMillimeters(OffsetZMaxMillimetersText, "Расширение Z вверх", out double offsetZMax, out errorMessage))
            {
                UserHelp = errorMessage;
                return false;
            }

            settings = new ZoomSettings
            {
                OffsetXMillimeters = offsetX,
                OffsetYMillimeters = offsetY,
                OffsetZMinMillimeters = offsetZMin,
                OffsetZMaxMillimeters = offsetZMax,
                FitCollision = FitCollision,
            };

            UserHelp = string.Empty;
            return true;
        }

        private void SetTextValue(ref string field, string value, [CallerMemberName] string propertyName = "")
        {
            if (field == value)
                return;

            field = value;
            MarkDirty();
            NotifyPropertyChanged(propertyName);
        }

        private void MarkDirty()
        {
            if (!_isApplyingConfig)
                IsDirty = true;

            UpdateCanSave();
        }

        private void UpdateCanSave()
        {
            CanSave = TryGetMillimeters(OffsetXMillimetersText, "Расширение X", out _, out _)
                && TryGetMillimeters(OffsetYMillimetersText, "Расширение Y", out _, out _)
                && TryGetMillimeters(OffsetZMinMillimetersText, "Расширение Z вниз", out _, out _)
                && TryGetMillimeters(OffsetZMaxMillimetersText, "Расширение Z вверх", out _, out _);
        }

        private static string FormatMillimeters(double value) => Math.Round(value).ToString(CultureInfo.CurrentCulture);

        private static bool TryGetMillimeters(string text, string name, out double value, out string errorMessage)
        {
            text = text?.Trim() ?? string.Empty;
            bool parsed = double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value)
                || double.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value);

            if (!parsed || value < 0)
            {
                errorMessage = $"Поле \"{name}\" должно быть числом больше или равно 0.";
                return false;
            }

            errorMessage = string.Empty;
            return true;
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}




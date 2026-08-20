using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_ModelChecker_User.Common
{
    public enum CheckListAnnotationsDisplayMode
    {
        SelectElements,
        HighlightElements
    }

    public sealed class CheckListAnnotationsSettingsM : IJsonSerializable, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private CheckListAnnotationsDisplayMode _displayMode = CheckListAnnotationsDisplayMode.SelectElements;
        private Color _highlightColor = new Color() { A = 255, B = 40, G = 60, R = 255 };
        private bool _isHighlightMode = false;
        private bool _clearPreviousHighlight = true;

        [JsonConstructor]
        public CheckListAnnotationsSettingsM() { }

        public CheckListAnnotationsDisplayMode DisplayMode
        {
            get => _displayMode;
            set
            {
                _displayMode = value;
                NotifyPropertyChanged();

                IsHighlightMode = _displayMode == CheckListAnnotationsDisplayMode.HighlightElements;
            }
        }

        public Color HighlightColor
        {
            get => _highlightColor;
            set
            {
                _highlightColor = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(HighlightColorBrush));
            }
        }

        public SolidColorBrush HighlightColorBrush => new SolidColorBrush(HighlightColor);

        public bool IsHighlightMode
        {
            get => _isHighlightMode;
            set
            {
                _isHighlightMode = value;
                NotifyPropertyChanged();
            }
        }

        public bool ClearPreviousHighlight
        {
            get => _clearPreviousHighlight;
            set
            {
                _clearPreviousHighlight = value;
                NotifyPropertyChanged();
            }
        }

        public object ToJson() => new
        {
            this.HighlightColor,
            this.ClearPreviousHighlight,
            this.DisplayMode,
        };

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_TrailingMEP.Forms.Entities
{
    /// <summary>
    /// Модель окна продолжения MEP трасс.
    /// </summary>
    public sealed class RouteTraceM : INotifyPropertyChanged
    {
        private const double MinRouteLength = 0.01;

        private readonly List<MepCurveData> _sourceCurves = new List<MepCurveData>();
        private readonly List<XYZ> _rawRoutePoints = new List<XYZ>();
        private readonly List<ElementId> _previewRouteIds = new List<ElementId>();
        private string _userMainStatus;
        private string _userHelp;
        private Brush _userMainStatusBrush = Brushes.Orange;
        private Brush _userHelpBrush = Brushes.Orange;
        private RouteMessageKind _currentMessageKind = RouteMessageKind.Instruction;
        private bool _isPreviewRouteChanged;
        private int _routeChangeTrackingSuspendCount;
        private bool _deletePreviewAfterBuild = true;
        private bool _autoCorrectRoute = true;
        private bool _allowAngle90 = true;
        private bool _allowAngle45 = true;
        private bool _allowAngle30 = true;

        public RouteTraceM(UIApplication uiapp)
        {
            UIApp = uiapp;
            Doc = uiapp.ActiveUIDocument.Document;
            UserMainStatus = "Выбери пучок, укажи точки траектории и запускай построение.";
            UserHelp = string.Empty;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public event EventHandler RouteSettingsChanged;

        public UIApplication UIApp { get; set; }

        public Document Doc { get; set; }

        public IReadOnlyList<MepCurveData> SourceCurves => _sourceCurves;

        public IReadOnlyList<XYZ> RawRoutePoints => _rawRoutePoints;

        public IReadOnlyList<ElementId> PreviewRouteIds => _previewRouteIds;

        public bool DeletePreviewAfterBuild
        {
            get => _deletePreviewAfterBuild;
            set
            {
                _deletePreviewAfterBuild = value;
                NotifyPropertyChanged();
            }
        }

        public bool AutoCorrectRoute
        {
            get => _autoCorrectRoute;
            set
            {
                if (_autoCorrectRoute == value)
                    return;

                _autoCorrectRoute = value;
                NotifyPropertyChanged();
                NotifyStateChanged();
                RouteSettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public bool AllowAngle90
        {
            get => _allowAngle90;
            set
            {
                if (_allowAngle90 == value)
                    return;

                _allowAngle90 = value;
                NotifyPropertyChanged();
                NotifyStateChanged();
                RouteSettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public bool AllowAngle45
        {
            get => _allowAngle45;
            set
            {
                if (_allowAngle45 == value)
                    return;

                _allowAngle45 = value;
                NotifyPropertyChanged();
                NotifyStateChanged();
                RouteSettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public bool AllowAngle30
        {
            get => _allowAngle30;
            set
            {
                if (_allowAngle30 == value)
                    return;

                _allowAngle30 = value;
                NotifyPropertyChanged();
                NotifyStateChanged();
                RouteSettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public string UserMainStatus
        {
            get => _userMainStatus;
            set
            {
                string status = value ?? string.Empty;
                _currentMessageKind = GetMessageKind(status);
                _userMainStatusBrush = GetMessageBrush(_currentMessageKind);
                _userHelpBrush = _currentMessageKind == RouteMessageKind.Error ? Brushes.IndianRed : Brushes.Orange;
                _userMainStatus = string.IsNullOrWhiteSpace(status) ? string.Empty : $"ВАЖНО: {status}";
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(UserMainStatusBrush));
                NotifyPropertyChanged(nameof(UserHelpBrush));
            }
        }

        public Brush UserMainStatusBrush => _userMainStatusBrush;

        public string UserHelp
        {
            get => _userHelp;
            set
            {
                _userHelp = value ?? string.Empty;
                _userHelpBrush = _currentMessageKind == RouteMessageKind.Error ? Brushes.IndianRed : Brushes.Orange;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(UserHelpBrush));
            }
        }

        public Brush UserHelpBrush => _userHelpBrush;

        public string BundleInfo => _sourceCurves.Count == 0
            ? "Пучок пока не выбран."
            : $"Выбрано элементов: {_sourceCurves.Count}.";

        public string TargetInfo => _rawRoutePoints.Count == 0
            ? "Точки траектории пока не указаны."
            : $"Точек траектории: {_rawRoutePoints.Count}. Z продолжений берется отдельно из каждого элемента пучка.";

        public string PreviewInfo => IsPreviewRouteChanged
            ? "Траектория изменилась. Выбери траекторию заново, чтобы принять правки."
            : HasPreviewRoute
            ? $"Сегментов траектории: {_previewRouteIds.Count}. После ручной правки линий выбери траекторию заново."
            : "Траектория пока не создана.";

        public string AngleInfo => AutoCorrectRoute
            ? $"Автокорректировка включена. Углы: {string.Join(", ", GetAllowedAngles().Select(a => a.ToString()))}."
            : "Автокорректировка выключена: сегменты строятся по кратчайшему пути между кликами.";

        public bool CanPickPoints => _sourceCurves.Count > 0;

        public bool CanCreatePreview => _sourceCurves.Count > 0 && _rawRoutePoints.Count > 0;

        public bool HasPreviewRoute => _previewRouteIds.Any(id => id != null && !id.Equals(ElementId.InvalidElementId));

        public bool IsPreviewRouteChanged => _isPreviewRouteChanged;

        public bool IsRouteChangeTrackingSuspended => _routeChangeTrackingSuspendCount > 0;

        public bool CanPickPreviewRoute => _sourceCurves.Count > 0;

        public bool CanBuild => _sourceCurves.Count > 0 && HasPreviewRoute && !IsPreviewRouteChanged;

        public void SetSourceCurves(IEnumerable<MepCurveData> sourceCurves)
        {
            _sourceCurves.Clear();

            if (sourceCurves != null)
                _sourceCurves.AddRange(sourceCurves);

            NotifyStateChanged();
        }

        public void AddRawRoutePoint(XYZ point)
        {
            if (point == null)
                return;

            if (_rawRoutePoints.Count == 0)
                RecalculateSourceCurvesByTarget(point);

            _rawRoutePoints.Add(point);
            NotifyStateChanged();
        }

        public void SetRawRoutePoints(IEnumerable<XYZ> routePoints)
        {
            _rawRoutePoints.Clear();

            if (routePoints != null)
                _rawRoutePoints.AddRange(routePoints.Where(p => p != null));

            _isPreviewRouteChanged = false;
            RecalculateSourceCurvesByTarget();
            NotifyStateChanged();
        }

        public void SetPreviewRouteIds(IEnumerable<ElementId> routeIds)
        {
            _previewRouteIds.Clear();

            if (routeIds != null)
                _previewRouteIds.AddRange(routeIds.Where(id => id != null && !id.Equals(ElementId.InvalidElementId)));

            _isPreviewRouteChanged = false;
            NotifyStateChanged();
        }

        public void ClearPreview()
        {
            _previewRouteIds.Clear();
            _isPreviewRouteChanged = false;
            NotifyStateChanged();
        }

        public void ClearRoute()
        {
            _rawRoutePoints.Clear();
            ClearPreview();
        }

        public void BeginInternalRouteChange()
        {
            _routeChangeTrackingSuspendCount++;
        }

        public void EndInternalRouteChange()
        {
            if (_routeChangeTrackingSuspendCount > 0)
                _routeChangeTrackingSuspendCount--;
        }

        public bool HasPreviewRouteChanges(ICollection<ElementId> modifiedIds, ICollection<ElementId> deletedIds)
        {
            if (!HasPreviewRoute || IsPreviewRouteChanged)
                return false;

            return HasAnyPreviewRouteId(modifiedIds) || HasAnyPreviewRouteId(deletedIds);
        }

        public void MarkPreviewRouteChanged()
        {
            if (!HasPreviewRoute || IsPreviewRouteChanged)
                return;

            _isPreviewRouteChanged = true;
            UserMainStatus = "Траектория изменилась.";
            UserHelp = "Выбери траекторию заново, чтобы принять новую геометрию перед построением.";
            NotifyStateChanged();
        }

        public XYZ GetBundleBasePoint()
        {
            if (_sourceCurves.Count == 0)
                return null;

            double x = _sourceCurves.Average(c => c.ExtensionStart.X);
            double y = _sourceCurves.Average(c => c.ExtensionStart.Y);
            double z = _sourceCurves.Average(c => c.ExtensionStart.Z);
            return new XYZ(x, y, z);
        }

        public XYZ GetBundleBaseDirection()
        {
            if (_sourceCurves.Count == 0)
                return XYZ.BasisX;

            XYZ summary = XYZ.Zero;
            foreach (MepCurveData sourceCurve in _sourceCurves)
            {
                XYZ direction = RouteBuilder.ProjectToXY(sourceCurve.SourceDirection);
                if (direction.GetLength() > 1e-9)
                    summary += direction.Normalize();
            }

            return summary.GetLength() > 1e-9 ? summary.Normalize() : XYZ.BasisX;
        }

        public IReadOnlyList<int> GetAllowedAngles()
        {
            List<int> angles = new List<int>();

            if (AllowAngle90)
                angles.Add(90);
            if (AllowAngle45)
                angles.Add(45);
            if (AllowAngle30)
                angles.Add(30);

            return angles;
        }

        public bool HasValidRouteData(out string reason)
        {
            reason = string.Empty;

            if (_sourceCurves.Count == 0)
            {
                reason = "Выбери хотя бы один MEP-элемент.";
                return false;
            }

            if (_rawRoutePoints.Count == 0)
            {
                reason = "Укажи хотя бы одну точку траектории.";
                return false;
            }

            XYZ basePoint = GetBundleBasePoint();
            if (basePoint == null || RouteBuilder.DistanceXY(basePoint, _rawRoutePoints.First()) < MinRouteLength)
            {
                reason = "Первая точка траектории слишком близко к торцу пучка.";
                return false;
            }

            if (AutoCorrectRoute && GetAllowedAngles().Count == 0)
            {
                reason = "Выбери хотя бы один угол автокорректировки или выключи автокорректировку.";
                return false;
            }

            return true;
        }

        public void NotifyStateChanged()
        {
            NotifyPropertyChanged(nameof(BundleInfo));
            NotifyPropertyChanged(nameof(TargetInfo));
            NotifyPropertyChanged(nameof(PreviewInfo));
            NotifyPropertyChanged(nameof(AngleInfo));
            NotifyPropertyChanged(nameof(CanPickPoints));
            NotifyPropertyChanged(nameof(CanCreatePreview));
            NotifyPropertyChanged(nameof(HasPreviewRoute));
            NotifyPropertyChanged(nameof(IsPreviewRouteChanged));
            NotifyPropertyChanged(nameof(CanPickPreviewRoute));
            NotifyPropertyChanged(nameof(CanBuild));
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

        private void RecalculateSourceCurvesByTarget()
        {
            if (_rawRoutePoints.Count == 0 || _sourceCurves.Count == 0)
                return;

            RecalculateSourceCurvesByTarget(_rawRoutePoints.First());
        }

        private void RecalculateSourceCurvesByTarget(XYZ targetPoint)
        {
            if (targetPoint == null || _sourceCurves.Count == 0)
                return;

            Element[] sourceElements = _sourceCurves
                .Select(c => Doc.GetElement(c.SourceId))
                .Where(e => e != null)
                .ToArray();

            _sourceCurves.Clear();
            _sourceCurves.AddRange(sourceElements.Select(e => RouteBuilder.CreateMepCurveData(Doc, e, targetPoint)));
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static RouteMessageKind GetMessageKind(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return RouteMessageKind.Instruction;

            if (message.StartsWith("Не удалось", StringComparison.OrdinalIgnoreCase)
                || message.IndexOf("ошиб", StringComparison.OrdinalIgnoreCase) >= 0)
                return RouteMessageKind.Error;

            if (message.StartsWith("Траектория изменилась", StringComparison.OrdinalIgnoreCase))
                return RouteMessageKind.Instruction;

            if (message.StartsWith("Выбрано", StringComparison.OrdinalIgnoreCase)
                || message.StartsWith("Построено", StringComparison.OrdinalIgnoreCase)
                || message.StartsWith("Обработано", StringComparison.OrdinalIgnoreCase)
                || message.StartsWith("Траектория", StringComparison.OrdinalIgnoreCase)
                || message.StartsWith("Точек траектории", StringComparison.OrdinalIgnoreCase))
                return RouteMessageKind.Success;

            return RouteMessageKind.Instruction;
        }

        private static Brush GetMessageBrush(RouteMessageKind messageKind)
        {
            switch (messageKind)
            {
                case RouteMessageKind.Success:
                    return Brushes.LightGreen;
                case RouteMessageKind.Error:
                    return Brushes.IndianRed;
                default:
                    return Brushes.Orange;
            }
        }

        private bool HasAnyPreviewRouteId(ICollection<ElementId> elementIds)
        {
            if (elementIds == null || elementIds.Count == 0)
                return false;

            return _previewRouteIds.Any(routeId => elementIds.Any(changedId => routeId.Equals(changedId)));
        }
    }

    internal enum RouteMessageKind
    {
        Instruction,
        Success,
        Error
    }
}

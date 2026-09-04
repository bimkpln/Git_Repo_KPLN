using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using KPLN_MEPBender.Common;
using KPLN_MEPBender.Services.Geometry;
using KPLN_MEPBender.Services.Routing;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_MEPBender.Forms.Entities
{
    public sealed class MepBenderM : INotifyPropertyChanged, IJsonSerializable
    {
        private double _offsetMm = 100;
        private double _offsetIterationStepMm = 25;
        private double _angleDegrees = 45;
        private bool _bendUp;
        private bool _bendDown;
        private bool _bendLeft = true;
        private bool _bendRight;
        private bool _analyzeCollisions;
        private bool _autoClearObstaclesAfterRun = true;
        private bool _autoClearRoutesAfterRun = true;
        private bool _alignVerticalBendByLowest = true;
        private string _userMainStatus;
        private string _userHelp;
        private string _userHelpForeground = "#FFD166";

        public MepBenderM()
        {
            ObstacleReferences = new ObservableCollection<LinkedElementReference>();
            RouteElementIds = new ObservableCollection<ElementId>();
            AngleOptions = new ObservableCollection<double>(BendAngleCatalog.GetUserAngles());
            UserHelp = "Выбери элементы-препятствия с 3D-геометрией, затем MEP-трассы для изменения.";
        }

        public ObservableCollection<LinkedElementReference> ObstacleReferences { get; }

        public ObservableCollection<ElementId> RouteElementIds { get; }

        public ObservableCollection<double> AngleOptions { get; }

        public double OffsetMm
        {
            get => _offsetMm;
            set
            {
                _offsetMm = value;
                OnPropertyChanged();
                UpdateCanRunProperties();
            }
        }

        public double OffsetIterationStepMm
        {
            get => _offsetIterationStepMm;
            set
            {
                _offsetIterationStepMm = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OffsetIterationHelp));
                UpdateCanRunProperties();
            }
        }

        public double AngleDegrees
        {
            get => _angleDegrees;
            set
            {
                _angleDegrees = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AngleHelp));
                UpdateCanRunProperties();
            }
        }

        public bool BendUp
        {
            get => _bendUp;
            set => SetDirection(BendDirection.Up, value);
        }

        public bool BendDown
        {
            get => _bendDown;
            set => SetDirection(BendDirection.Down, value);
        }

        public bool BendLeft
        {
            get => _bendLeft;
            set => SetDirection(BendDirection.Left, value);
        }

        public bool BendRight
        {
            get => _bendRight;
            set => SetDirection(BendDirection.Right, value);
        }

        public bool AnalyzeCollisions
        {
            get => _analyzeCollisions;
            set
            {
                _analyzeCollisions = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ClashAnalyzeHelp));
            }
        }

        public bool AutoClearObstaclesAfterRun
        {
            get => _autoClearObstaclesAfterRun;
            set
            {
                _autoClearObstaclesAfterRun = value;
                OnPropertyChanged();
            }
        }

        public bool AutoClearRoutesAfterRun
        {
            get => _autoClearRoutesAfterRun;
            set
            {
                _autoClearRoutesAfterRun = value;
                OnPropertyChanged();
            }
        }

        public bool AlignVerticalBendByLowest
        {
            get => _alignVerticalBendByLowest;
            set
            {
                _alignVerticalBendByLowest = value;
                OnPropertyChanged();
            }
        }

        public string UserMainStatus
        {
            get => _userMainStatus;
            set
            {
                _userMainStatus = value;
                OnPropertyChanged();
            }
        }

        public string UserHelp
        {
            get => _userHelp;
            set
            {
                _userHelp = value;
                OnPropertyChanged();
            }
        }

        public string UserHelpForeground
        {
            get => _userHelpForeground;
            set
            {
                _userHelpForeground = value;
                OnPropertyChanged();
            }
        }

        public string ObstaclesButtonText => $"Огибаемые элементы ({ObstacleReferences.Count})";

        public string RoutesButtonText => $"Выбрать трассы ({RouteElementIds.Count})";

        public string AngleHelp => $"Угол построения огибания: {AngleDegrees}°.";

        public string OffsetIterationHelp => $"Добор зазора применяется только если Revit сообщает, что места не хватило. Максимум 5 пересчётов: от минимального зазора до +{OffsetIterationStepMm * 5:0.##} мм.";

        public string AlignVerticalBendHelp => "Для направления вниз выравнивает все выбранные параллельные трассы по самой нижней рассчитанной отметке обхода.";
        public string DirectionHelp => "Одновременно может быть активно только одно направление. Плоскость изменения определит алгоритм.";

        public string ClashAnalyzeHelp => AnalyzeCollisions
            ? "Подключение к KPLN_IOSClasher заложено в отдельном сервисе, сам анализ пока не запускается."
            : "Анализ коллизий можно будет включить после подключения KPLN_IOSClasher.";

        public bool CanRun => ObstacleReferences.Count > 0
                              && RouteElementIds.Count > 0
                              && GetSelectedDirections().Any()
                              && OffsetMm > 0
                              && OffsetIterationStepMm > 0
                              && AngleOptions.Contains(AngleDegrees);

        public IReadOnlyCollection<BendDirection> GetSelectedDirections()
        {
            List<BendDirection> directions = new List<BendDirection>();

            if (BendUp)
                directions.Add(BendDirection.Up);
            if (BendDown)
                directions.Add(BendDirection.Down);
            if (BendLeft)
                directions.Add(BendDirection.Left);
            if (BendRight)
                directions.Add(BendDirection.Right);

            return directions;
        }

        public object ToJson()
        {
            return new
            {
                OffsetMm,
                OffsetIterationStepMm,
                AngleDegrees,
                BendUp,
                BendDown,
                BendLeft,
                BendRight,
                AlignVerticalBendByLowest,
                AnalyzeCollisions,
                AutoClearObstaclesAfterRun,
                AutoClearRoutesAfterRun
            };
        }

        public void NormalizeConfigValues()
        {
            if (!AngleOptions.Contains(AngleDegrees))
                AngleDegrees = 45;

            if (OffsetMm <= 0)
                OffsetMm = 100;

            if (OffsetIterationStepMm <= 0)
                OffsetIterationStepMm = 25;

            NormalizeDirectionSelection();
        }

        public void SetObstacles(IEnumerable<LinkedElementReference> references)
        {
            ResetCollection(ObstacleReferences, references);
            OnPropertyChanged(nameof(ObstaclesButtonText));
            UpdateCanRunProperties();
        }

        public void AddObstacles(IEnumerable<LinkedElementReference> references)
        {
            HashSet<string> existingKeys = new HashSet<string>(ObstacleReferences.Select(GetReferenceKey));
            foreach (LinkedElementReference reference in references)
            {
                string key = GetReferenceKey(reference);
                if (!existingKeys.Add(key))
                    continue;

                ObstacleReferences.Add(reference);
            }

            OnPropertyChanged(nameof(ObstaclesButtonText));
            UpdateCanRunProperties();
        }

        public void SetRoutes(IEnumerable<ElementId> ids)
        {
            ResetCollection(RouteElementIds, ids);
            OnPropertyChanged(nameof(RoutesButtonText));
            UpdateCanRunProperties();
        }

        public void ClearObstacles()
        {
            ObstacleReferences.Clear();
            OnPropertyChanged(nameof(ObstaclesButtonText));
            UpdateCanRunProperties();
        }

        public void ClearRoutes()
        {
            RouteElementIds.Clear();
            OnPropertyChanged(nameof(RoutesButtonText));
            UpdateCanRunProperties();
        }

        public void SetStatus(string mainStatus, string help = null, string helpForeground = null)
        {
            UserMainStatus = mainStatus;
            if (help != null)
                UserHelp = help;
            if (!string.IsNullOrWhiteSpace(helpForeground))
                UserHelpForeground = helpForeground;
        }

        private void SetDirection(BendDirection direction, bool isSelected)
        {
            bool oldBendUp = _bendUp;
            bool oldBendDown = _bendDown;
            bool oldBendLeft = _bendLeft;
            bool oldBendRight = _bendRight;

            if (isSelected)
            {
                _bendUp = direction == BendDirection.Up;
                _bendDown = direction == BendDirection.Down;
                _bendLeft = direction == BendDirection.Left;
                _bendRight = direction == BendDirection.Right;
            }
            else
            {
                switch (direction)
                {
                    case BendDirection.Up:
                        _bendUp = false;
                        break;
                    case BendDirection.Down:
                        _bendDown = false;
                        break;
                    case BendDirection.Left:
                        _bendLeft = false;
                        break;
                    case BendDirection.Right:
                        _bendRight = false;
                        break;
                }
            }

            RaiseDirectionChanges(oldBendUp, oldBendDown, oldBendLeft, oldBendRight);
        }

        private void NormalizeDirectionSelection()
        {
            if (_bendLeft)
                SetDirection(BendDirection.Left, true);
            else if (_bendRight)
                SetDirection(BendDirection.Right, true);
            else if (_bendUp)
                SetDirection(BendDirection.Up, true);
            else if (_bendDown)
                SetDirection(BendDirection.Down, true);
            else
                SetDirection(BendDirection.Left, true);
        }

        private void RaiseDirectionChanges(bool oldBendUp, bool oldBendDown, bool oldBendLeft, bool oldBendRight)
        {
            if (oldBendUp != _bendUp)
                OnPropertyChanged(nameof(BendUp));
            if (oldBendDown != _bendDown)
                OnPropertyChanged(nameof(BendDown));
            if (oldBendLeft != _bendLeft)
                OnPropertyChanged(nameof(BendLeft));
            if (oldBendRight != _bendRight)
                OnPropertyChanged(nameof(BendRight));

            UpdateCanRunProperties();
        }

        private void ResetCollection(ObservableCollection<ElementId> collection, IEnumerable<ElementId> ids)
        {
            collection.Clear();
            foreach (ElementId id in ids.Distinct())
                collection.Add(id);
        }

        private void ResetCollection(ObservableCollection<LinkedElementReference> collection, IEnumerable<LinkedElementReference> references)
        {
            collection.Clear();
            foreach (LinkedElementReference reference in references.GroupBy(GetReferenceKey).Select(g => g.First()))
                collection.Add(reference);
        }

        private string GetReferenceKey(LinkedElementReference reference)
        {
            return $"{reference.HostElementId.GetStableIntegerValue()}:{reference.LinkedElementId.GetStableIntegerValue()}";
        }

        private void UpdateCanRunProperties()
        {
            OnPropertyChanged(nameof(CanRun));
            OnPropertyChanged(nameof(DirectionHelp));
            OnPropertyChanged(nameof(ClashAnalyzeHelp));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

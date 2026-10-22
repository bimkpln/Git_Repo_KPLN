using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_Tools_Lib.Forms.Entities
{
    internal sealed class TagSpacesM : INotifyPropertyChanged
    {
        private readonly TagSpacesTypeM[] _allTypes;
        private string _selectedCategory;
        private TagSpacesTypeM _selectedType;
        private TagSpacesTypeM _roomType;
        private TagSpacesTypeM _spaceType;
        private bool _useRadiators;
        private bool _useGrilles;
        private TagSpacesTypeM _radiatorType;
        private TagSpacesTypeM _grilleType;
        private TagSpacesConfig _loadedConfig;

        internal TagSpacesM(Document document, ViewPlan[] views)
        {
            HasSpaces = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType().Cast<Space>().Any(space => space.Area > 0);
            Categories = HasSpaces ? new[] { "Помещения", "Пространства" } : new[] { "Помещения" };

            Views = views;
            ViewNames = string.Join("\n", views.Select(view => view.Name));
            _allTypes = new FilteredElementCollector(document).OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>().Where(type => type is RoomTagType || type is SpaceTagType)
                .Select(type => new TagSpacesTypeM(type)).OrderBy(type => type.Name).ToArray();

            Sources = new[] { new TagSpacesSourceM(null) }.Concat(
                new FilteredElementCollector(document).OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>().Where(link => link.GetLinkDocument() != null)
                    .OrderBy(link => link.Name).Select(link => new TagSpacesSourceM(link))).ToArray();

            foreach (TagSpacesSourceM source in Sources)
                source.PropertyChanged += (sender, args) => UpdateValidation();

            (Sources.Length > 1 ? Sources[1] : Sources[0]).IsSelected = true;
            _roomType = _allTypes.FirstOrDefault(type => type.IsRoom);
            _spaceType = _allTypes.Where(type => !type.IsRoom)
                .OrderByDescending(type => type.RevitType.FamilyName.IndexOf("Пространство_Имя_Номер_Площадь", System.StringComparison.OrdinalIgnoreCase) >= 0)
                .FirstOrDefault(type => new[] { "имя", "номер", "площадь" }.All(part =>
                    type.RevitType.Name.IndexOf(part, System.StringComparison.OrdinalIgnoreCase) >= 0))
                ?? _allTypes.FirstOrDefault(type => !type.IsRoom);
            SpaceTagTypes = _allTypes.Where(type => !type.IsRoom).ToArray();
            SelectedRadiatorType = FindDefaultSpaceTag("теплопот");
            SelectedGrilleType = FindDefaultSpaceTag("воздухообм");
            SelectedCategory = Categories[0];
        }

        internal ViewPlan[] Views { get; }
        public string ViewNames { get; }
        public string ViewSummary => string.Format("Выбрано планов: {0}", Views.Length);
        public string[] Categories { get; }
        public bool HasSpaces { get; }
        public TagSpacesSourceM[] Sources { get; }
        public TagSpacesTypeM[] TagTypes { get; private set; }
        public TagSpacesTypeM[] SpaceTagTypes { get; }
        public bool IsRoom => SelectedCategory == Categories[0];
        public bool IsSpace => !IsRoom;
        public string MainTagTypeLabel => IsRoom ? "Тип марки для всех помещений" : "Тип марки для всех пространств";
        public bool CanRun => string.IsNullOrEmpty(ValidationError);
        public string ConfigStatus { get; internal set; }
        public string UserHelp => CanRun
            ? "Уже маркированные элементы пропускаются. При обрезке аннотаций скрытая марка заменяется новой, если новая видна на плане."
            : ValidationError;

        private string ValidationError
        {
            get
            {
                if (Views.Length == 0)
                    return "Выберите планы для маркировки.";
                if (SelectedType == null || SelectedType.IsRoom != IsRoom)
                    return "Выберите загруженный тип марки для выбранной категории.";
                if (IsRoom && !Sources.Any(source => source.IsSelected))
                    return "Отметьте хотя бы один источник помещений. Если нужной связи нет в списке, загрузите её и откройте команду заново.";
                if (IsSpace && UseRadiators && SelectedRadiatorType == null)
                    return "Выберите тип марки пространства с радиатором.";
                if (IsSpace && UseGrilles && SelectedGrilleType == null)
                    return "Выберите тип марки пространства с решёткой.";
                return string.Empty;
            }
        }

        public bool UseRadiators
        {
            get => _useRadiators;
            set
            {
                _useRadiators = value;
                NotifyPropertyChanged();
                UpdateValidation();
            }
        }

        public bool UseGrilles
        {
            get => _useGrilles;
            set
            {
                _useGrilles = value;
                NotifyPropertyChanged();
                UpdateValidation();
            }
        }

        public TagSpacesTypeM SelectedRadiatorType
        {
            get => _radiatorType;
            set
            {
                _radiatorType = value;
                NotifyPropertyChanged();
                UpdateValidation();
            }
        }

        public TagSpacesTypeM SelectedGrilleType
        {
            get => _grilleType;
            set
            {
                _grilleType = value;
                NotifyPropertyChanged();
                UpdateValidation();
            }
        }

        public string SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                _selectedCategory = value;
                TagTypes = _allTypes.Where(type => type.IsRoom == IsRoom).ToArray();
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(IsRoom));
                NotifyPropertyChanged(nameof(IsSpace));
                NotifyPropertyChanged(nameof(MainTagTypeLabel));
                NotifyPropertyChanged(nameof(TagTypes));

                // Сначала обновляем ItemsSource, затем выбор в новом списке типов.
                SelectedType = IsRoom ? _roomType : _spaceType;
                UpdateValidation();
            }
        }

        public TagSpacesTypeM SelectedType
        {
            get => _selectedType;
            set
            {
                _selectedType = value;
                if (value != null && value.IsRoom == IsRoom)
                {
                    if (IsRoom)
                        _roomType = value;
                    else
                        _spaceType = value;
                }
                NotifyPropertyChanged();
                UpdateValidation();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        internal void ApplyConfig(TagSpacesConfig settings)
        {
            if (settings == null)
                return;

            _loadedConfig = settings;
            _roomType = ResolveType(settings.RoomTag, true) ?? (settings.RoomTag == null ? _roomType : null);
            _spaceType = ResolveType(settings.SpaceTag, false) ?? (settings.SpaceTag == null ? _spaceType : null);
            TagSpacesSourceConfig[] savedSources = settings.Sources ?? new[]
            {
                new TagSpacesSourceConfig
                {
                    LinkUniqueId = settings.UseLinkedRooms ? settings.LinkUniqueId : null,
                    LinkName = settings.UseLinkedRooms ? settings.LinkName : null
                }
            };
            foreach (TagSpacesSourceM source in Sources)
                source.IsSelected = false;

            foreach (TagSpacesSourceConfig saved in savedSources)
            {
                if (saved == null)
                    continue;

                TagSpacesSourceM source = string.IsNullOrEmpty(saved.LinkUniqueId) && string.IsNullOrEmpty(saved.LinkName)
                    ? Sources[0]
                    : Sources.FirstOrDefault(item => item.RevitLink != null && item.RevitLink.UniqueId == saved.LinkUniqueId)
                        ?? Sources.FirstOrDefault(item => item.RevitLink != null && item.RevitLink.Name == saved.LinkName);
                if (source != null)
                    source.IsSelected = true;
            }

            SelectedRadiatorType = ResolveType(settings.RadiatorTag, false) ?? FindDefaultSpaceTag("теплопот");
            SelectedGrilleType = ResolveType(settings.GrilleTag, false) ?? FindDefaultSpaceTag("воздухообм");
            UseRadiators = settings.UseRadiators;
            UseGrilles = settings.UseGrilles;
            SelectedCategory = settings.IsRoom || !HasSpaces ? Categories[0] : Categories[1];
        }

        internal TagSpacesConfig GetConfig()
        {
            return new TagSpacesConfig
            {
                IsRoom = IsRoom,
                RoomTag = ToConfig(_roomType) ?? _loadedConfig?.RoomTag,
                SpaceTag = ToConfig(_spaceType) ?? _loadedConfig?.SpaceTag,
                Sources = IsRoom || _loadedConfig == null
                    ? Sources.Where(source => source.IsSelected).Select(source => new TagSpacesSourceConfig
                    {
                        LinkUniqueId = source.RevitLink?.UniqueId,
                        LinkName = source.RevitLink?.Name
                    }).ToArray()
                    : _loadedConfig.Sources,
                UseLinkedRooms = _loadedConfig?.UseLinkedRooms == true,
                LinkUniqueId = _loadedConfig?.LinkUniqueId,
                LinkName = _loadedConfig?.LinkName,
                UseRadiators = UseRadiators,
                UseGrilles = UseGrilles,
                RadiatorTag = ToConfig(SelectedRadiatorType) ?? _loadedConfig?.RadiatorTag,
                GrilleTag = ToConfig(SelectedGrilleType) ?? _loadedConfig?.GrilleTag
            };
        }

        private TagSpacesTypeM FindDefaultSpaceTag(string namePart)
            => SpaceTagTypes.FirstOrDefault(type => type.Name.IndexOf(namePart, System.StringComparison.OrdinalIgnoreCase) >= 0);

        private TagSpacesTypeM ResolveType(TagSpacesTypeConfig saved, bool isRoom)
            => saved == null ? null : _allTypes.FirstOrDefault(type => type.IsRoom == isRoom
                && type.RevitType.FamilyName == saved.FamilyName && type.RevitType.Name == saved.TypeName);

        private static TagSpacesTypeConfig ToConfig(TagSpacesTypeM type)
            => type == null ? null : new TagSpacesTypeConfig
            {
                FamilyName = type.RevitType.FamilyName,
                TypeName = type.RevitType.Name
            };

        private void UpdateValidation()
        {
            NotifyPropertyChanged(nameof(CanRun));
            NotifyPropertyChanged(nameof(UserHelp));
            CommandManager.InvalidateRequerySuggested();
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    internal sealed class TagSpacesTypeM
    {
        internal TagSpacesTypeM(FamilySymbol type)
        {
            RevitType = type;
            IsRoom = type is RoomTagType;
            Name = string.Format("{0}: {1}", type.FamilyName, type.Name);
        }

        internal FamilySymbol RevitType { get; }
        internal bool IsRoom { get; }
        public string Name { get; }
    }

    internal sealed class TagSpacesSourceM : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                    return;

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        internal TagSpacesSourceM(RevitLinkInstance link)
        {
            RevitLink = link;
            Name = link == null ? "Текущая модель" : string.Format("{0} [id: {1}]", link.Name, link.Id);
        }

        internal RevitLinkInstance RevitLink { get; }
        public string Name { get; }
    }
}

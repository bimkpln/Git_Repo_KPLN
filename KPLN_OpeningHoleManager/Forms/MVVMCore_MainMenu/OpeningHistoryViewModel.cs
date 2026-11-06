using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Forms.MVVMCommand;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    public sealed class OpeningHistoryViewModel : INotifyPropertyChanged
    {
        private Document _document;
        private int _generation;
        private bool _queued;
        private bool _cacheMissing;
        private bool _suspended;
        private bool _fullRefresh;
        private bool _refreshLevels;
        private bool _refreshVisibility;
        private readonly Dictionary<string, ElementId> _pendingElements = new Dictionary<string, ElementId>();
        private readonly OpeningHistoryChangeTracker _tracker = new OpeningHistoryChangeTracker();
        private ObservableCollection<OpeningHistoryRow> _cards = new ObservableCollection<OpeningHistoryRow>();
        private readonly Dictionary<string, OpeningHistoryRow> _cardsById = new Dictionary<string, OpeningHistoryRow>();


        private readonly HashSet<string> _pendingCacheElementIds = new HashSet<string>();
        private readonly List<DocumentCreations> _documentCreations = new List<DocumentCreations>();
        private Dictionary<string, OpeningCreationInfo> _observedCreations = new Dictionary<string, OpeningCreationInfo>();


        private bool _accessChecked;
        private bool _isBimUser;
        private bool _confirmAllQueued;
        private string _actionStatus = string.Empty;
        private Dictionary<string, string> _confirmationNotes = new Dictionary<string, string>();


        private string _filter = string.Empty;
        private string _status = "Откройте модель и обновите список";
        private OpeningHistoryRow[] _rows = new OpeningHistoryRow[0];
        private HashSet<string> _elementIds = new HashSet<string>();
        private ICollectionView _items;
        private OpeningHistoryFilterSettings _filterSettings = new OpeningHistoryFilterSettings();
        private HashSet<string> _visibleIds;
        private string _viewFilterNote;
        private OpeningFilterOption[] _levels;
        private bool _onlySelected;
        private HashSet<string> _selectedIds = new HashSet<string>();
        private string _selectionError = string.Empty;


        public event PropertyChangedEventHandler PropertyChanged;
        public ICommand RefreshCommand { get; }
        public ICommand ShowCommand { get; }
        public ICommand SelectCommand { get; }
        public ICommand ConfirmCommand { get; }
        public ICommand ConfirmAllCommand { get; }


        public bool IsBimUser
        {
            get => _isBimUser;
            private set
            {
                _isBimUser = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }
        public string ActionStatus
        {
            get => _actionStatus;
            private set
            {
                _actionStatus = value;
                OnPropertyChanged();
            }
        }

        public ICollectionView Items
        {
            get => _items;
            private set
            {
                _items = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get => _status;
            private set
            {
                _status = value;
                OnPropertyChanged();
            }
        }

        public string Filter
        {
            get => _filter;
            set
            {
                _filter = value ?? string.Empty;
                Items?.Refresh();
                OnPropertyChanged();
                UpdateCounts();
            }
        }
        public string CountText => "Показано: " + (Items?.Cast<object>().Count() ?? 0) + " / " + _rows.Length;
        public string UnconfirmedCountText => "Не подтверждено: "
            + (Items?.Cast<OpeningHistoryRow>().Count(row => !row.IsConfirmed) ?? 0)
            + " / " + _rows.Count(row => !row.IsConfirmed);
        public bool HasUnconfirmed => _rows.Any(row => !row.IsConfirmed);
        public bool OnlyActiveView => _filterSettings.OnlyActiveView;
        public bool OnlySelected
        {
            get => _onlySelected;
            set
            {
                if (_onlySelected == value)
                    return;

                _onlySelected = value;
                _selectedIds.Clear();
                _selectionError = string.Empty;
                Items?.Refresh();
                OnPropertyChanged();
                UpdateCounts();
            }
        }
        public string SelectionFilterNote => !OnlySelected ? string.Empty
            : !string.IsNullOrEmpty(_selectionError) ? _selectionError
            : !_elementIds.Overlaps(_selectedIds) ? "Выделите отверстия в модели или отключите «Выбранные»." : string.Empty;
        public string FilterButtonText => _filterSettings.ActiveCount == 0 ? "Фильтры" : "Фильтры (" + _filterSettings.ActiveCount + ")";

        // Вызывается из Idling: чтение выделения доступно и в Revit 2020.
        // Обновляем только представление списка, без сбора геометрии и записи истории.
        internal void RefreshSelection(UIDocument uiDocument)
        {
            if (!OnlySelected || !string.IsNullOrEmpty(_selectionError))
                return;

            try
            {
                Document activeDocument = uiDocument?.Document;
                var selectedIds = _document != null && _document.IsValidObject
                    && activeDocument != null && activeDocument.IsValidObject && _document.Equals(activeDocument)
                    ? new HashSet<string>(uiDocument.Selection.GetElementIds().Select(id => id.ToString()))
                    : new HashSet<string>();
                if (_selectedIds.SetEquals(selectedIds))
                    return;

                _selectedIds = selectedIds;
            }
            catch (Exception ex)
            {
                _selectedIds.Clear();
                _selectionError = "Не удалось прочитать выделение. Выключите и включите «Выбранные». "
                    + OpeningUserMessage.FromException(ex);
            }

            Items?.Refresh();
            UpdateCounts();
        }

        internal OpeningHistoryFilterSettings CreateFilterSettings()
        {
            var settings = _filterSettings.Copy();
            var levels = _levels ?? _rows.Where(row => !string.IsNullOrEmpty(row.LevelId)).GroupBy(row => row.LevelId)
                    .OrderBy(group => group.First().LevelElevation)
                    .Select(group => new OpeningFilterOption { Id = group.Key, Name = group.First().LevelName }).ToArray();
            settings.Levels = new[]
                {
                    new OpeningFilterOption { Id = OpeningHistoryFilterSettings.AllLevels, Name = "Все уровни" },
                    new OpeningFilterOption { Id = string.Empty, Name = "Без уровня" }
                }.Concat(levels)
                .ToArray();
            return settings;
        }

        internal void ApplyFilterSettings(OpeningHistoryFilterSettings settings)
        {
            _filterSettings = settings.Copy();
            _visibleIds = null;
            Items?.Refresh();
            OnPropertyChanged(nameof(FilterButtonText));
            UpdateCounts();
            RequestVisibilityRefresh();
        }

        public OpeningHistoryViewModel()
        {
            Items = CollectionViewSource.GetDefaultView(_cards);
            Items.Filter = item => (!OnlySelected || _selectedIds.Contains(((OpeningHistoryRow)item).ElementId))
                && _filterSettings.Matches((OpeningHistoryRow)item, Filter, _visibleIds);


            RefreshCommand = new RelayCommand(() => RequestRefresh(_document, true));
            ShowCommand = new RelayCommandWithParams(row => QueueElement(row as OpeningHistoryRow, ElementAction.Show));
            SelectCommand = new RelayCommandWithParams(row => QueueElement(row as OpeningHistoryRow, ElementAction.Select));
            ConfirmCommand = new RelayCommandWithParams(row => QueueElement(row as OpeningHistoryRow, ElementAction.Confirm));
            ConfirmAllCommand = new RelayCommand(QueueConfirmAll,
                () => !_suspended && IsBimUser && _document != null && _document.IsValidObject && _rows.Length > 0 && !_confirmAllQueued);
        }

        internal bool ContainsElementId(ElementId id) => _elementIds.Contains(id.ToString());

        internal void SetDocument(Document document)
        {
            if (document != null && !document.IsValidObject)
                document = null;

            // Закрытый документ нельзя передавать даже аргументом в Document.Equals.
            if (_document == null && document == null
                || _document != null && _document.IsValidObject && document != null && _document.Equals(document))
                return;


            _document = document;
            CancelPending();
            _observedCreations = GetDocumentCreations(document);
            _tracker.Clear();

            _accessChecked = false;
            _confirmAllQueued = false;
            IsBimUser = false;
            ActionStatus = string.Empty;
            _confirmationNotes.Clear();
            _filterSettings = new OpeningHistoryFilterSettings();
            _visibleIds = null;
            _levels = null;
            _onlySelected = false;
            _selectedIds.Clear();
            _selectionError = string.Empty;
            OnPropertyChanged(nameof(OnlySelected));
            OnPropertyChanged(nameof(FilterButtonText));
            SetRows(new OpeningHistoryRow[0]);
            Status = document == null ? "Откройте проект Revit" : "Нажмите «Обновить» для загрузки списка";
        }

        internal void Resume()
        {
            _suspended = false;
        }

        internal void Suspend()
        {
            _suspended = true;
            CancelPending();
        }

        private void CancelPending()
        {
            ++_generation;
            _queued = false;
            _confirmAllQueued = false;
            ClearPendingRefresh();
        }

        private Dictionary<string, OpeningCreationInfo> GetDocumentCreations(Document document)
        {
            // Отмена очереди не отменяет уже полученные дату и автора.
            // Храним незаписанные сведения отдельно для каждой открытой модели.
            _documentCreations.RemoveAll(item => !item.Document.IsValidObject || item.Entries.Count == 0);
            if (document == null)
                return new Dictionary<string, OpeningCreationInfo>();

            DocumentCreations observations = _documentCreations.FirstOrDefault(item => item.Document.Equals(document));
            if (observations == null)
            {
                observations = new DocumentCreations(document);
                _documentCreations.Add(observations);
            }

            return observations.Entries;
        }

        private void ClearPendingRefresh()
        {
            _cacheMissing = false;
            _fullRefresh = false;
            _refreshLevels = false;
            _refreshVisibility = false;
            _pendingElements.Clear();
            _pendingCacheElementIds.Clear();
        }


        internal void HandleDocumentChanged(Document document, DocumentChangedEventArgs args)
        {
            if (_suspended)
                return;

            OpeningHistoryChanges changes = _tracker.Read(document, args, OnlyActiveView);
            if (changes.CancelAutoCache)
            {
                // Undo/Redo отменяют запись, но сохраняют дату и автора для последующего повтора.
                _cacheMissing = false;
                _pendingCacheElementIds.Clear();
            }

            if (!changes.HasChanges)
                return;


            foreach (FamilyInstance opening in changes.Created)
                ObserveCreation(opening);

            _fullRefresh |= changes.FullRefresh;
            _refreshLevels |= changes.RefreshLevels;
            _refreshVisibility |= changes.RefreshVisibility;
            RequestElements(changes.Elements.Values);
        }

        private void ObserveCreation(Element element)
        {
            // Фиксируем дату и автора в DocumentChanged, а не в момент выполнения очереди.
            _pendingCacheElementIds.Add(element.Id.ToString());
            if (!_observedCreations.ContainsKey(element.UniqueId))
                _observedCreations.Add(element.UniqueId, OpeningHistoryService.CaptureCreation(element));
        }

        internal void RequestRefresh(Document document, bool cacheMissing, IEnumerable<ElementId> cacheElementIds = null, bool cancelAutoCache = false)
        {
            SetDocument(document);
            if (_suspended || document == null || !document.IsValidObject)
                return;

            _fullRefresh = true;
            _cacheMissing |= cacheMissing;

            if (cancelAutoCache)
            {
                _cacheMissing = false;
                _pendingCacheElementIds.Clear();
            }
            else if (cacheElementIds != null)
            {
                foreach (ElementId id in cacheElementIds)
                {
                    Element element = document.GetElement(id);
                    if (OpeningHistoryService.IsOpening(element))
                        ObserveCreation(element);
                }
            }

            EnqueueRefresh();
        }

        internal void RequestVisibilityRefresh()
        {
            _refreshVisibility = true;
            EnqueueRefresh();
        }

        private void RequestElements(IEnumerable<ElementId> ids)
        {
            foreach (ElementId id in ids.Where(OpeningHistoryService.IsValidId))
                _pendingElements[id.ToString()] = id;

            EnqueueRefresh();
        }

        private void EnqueueRefresh()
        {
            if (_suspended || _queued || _document == null || !_document.IsValidObject)
                return;

            _queued = true;
            int generation = _generation;
            Document document = _document;

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new HistoryCommand(document, app =>
            {
                if (generation != _generation)
                    return;

                bool full = _fullRefresh;
                bool write = _cacheMissing;
                bool levels = _refreshLevels;
                bool visibility = _refreshVisibility;
                ElementId[] ids = _pendingElements.Values.ToArray();
                var cacheIds = new HashSet<string>(_pendingCacheElementIds);

                // Несколько транзакций до Idling объединяются в одно обновление.
                _queued = false;
                ClearPendingRefresh();
                Refresh(document, full, write, levels, visibility, ids, cacheIds);
            }, message =>
            {
                if (generation != _generation)
                    return;

                _queued = false;
                ClearPendingRefresh();
                Status = message;
            }));
        }


        private void Refresh(Document document, bool full, bool cacheMissing, bool levels, bool visibility,
            ElementId[] ids, ISet<string> cacheElementIds)
        {
            if (!_accessChecked || cacheMissing)
            {
                IsBimUser = OpeningConfirmationService.IsBimUser();
                _accessChecked = true;
            }

            if (full || levels)
                RefreshLevels(document);

            if (full || visibility)
                RefreshVisibleIds(document);


            if (full || ids.Length > 0)
            {
                OpeningHistoryRow[] rows = OpeningHistoryService.Collect(document, cacheMissing, cacheElementIds,
                    _observedCreations, full ? null : ids);

                if (full)
                    SetRows(rows);
                else
                    ApplyRows(ids, rows);

                // Убираем сведения из памяти только после чтения карточки без ошибок записи.
                // Отсутствующее отверстие может вернуться через Undo/Redo с тем же UniqueId.
                foreach (OpeningHistoryRow row in rows.Where(row => string.IsNullOrEmpty(row.CacheNote)))
                    _observedCreations.Remove(row.UniqueId);
            }

            if (full || levels || visibility)
                Items.Refresh();

            UpdateCounts();
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            int warnings = _rows.Count(row => !string.IsNullOrEmpty(row.CacheNote));
            Status = _rows.Length == 0 ? "В модели нет поддерживаемых отверстий"
                : warnings > 0 ? "Есть замечания к записи истории: " + warnings : "История загружена из модели";

            if (!string.IsNullOrEmpty(_viewFilterNote))
                Status = _viewFilterNote;
        }


        private void RefreshLevels(Document document)
        {
            _levels = new FilteredElementCollector(document).OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(level => level.Elevation).ThenBy(level => level.Name)
                .Select(level => new OpeningFilterOption { Id = level.Id.ToString(), Name = level.Name }).ToArray();

            _tracker.SetLevels(_levels.Select(level => level.Id));

            // Пустой результат не сбрасывает выбранный уровень. Сбрасываем только удалённый уровень.
            if (!string.IsNullOrEmpty(_filterSettings.LevelId) && _filterSettings.LevelId != OpeningHistoryFilterSettings.AllLevels
                && !_levels.Any(level => level.Id == _filterSettings.LevelId))
            {
                _filterSettings.LevelId = OpeningHistoryFilterSettings.AllLevels;
                OnPropertyChanged(nameof(FilterButtonText));
            }
        }

        private void RefreshVisibleIds(Document document)
        {
            _visibleIds = null;
            _viewFilterNote = string.Empty;
            _tracker.SetViewDependencies(new ElementId[0]);

            if (!OnlyActiveView)
                return;

            try
            {
                View view = document.ActiveView;
                if (view == null || !FilteredElementCollector.IsViewValidForElementIteration(document, view.Id))
                {
                    _viewFilterNote = "На этом виде нельзя собрать отверстия. Откройте план, разрез или 3D-вид либо отключите фильтр активного вида.";
                    return;
                }

                // Настройки вида/шаблона/фильтров меняют видимость, но не историю отверстий.
                var dependencies = new List<ElementId> { view.Id, view.ViewTemplateId };
                dependencies.AddRange(view.GetDependentElements(null));
                if (view.AreGraphicsOverridesAllowed())
                    dependencies.AddRange(view.GetFilters());

                View template = document.GetElement(view.ViewTemplateId) as View;
                if (template != null && template.AreGraphicsOverridesAllowed())
                    dependencies.AddRange(template.GetFilters());

                dependencies.Add(view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId());
                dependencies.Add(view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)?.AsElementId());
                _tracker.SetViewDependencies(dependencies);

                _visibleIds = new HashSet<string>(new FilteredElementCollector(document, view.Id)
                    .WherePasses(OpeningHistoryService.CandidateFilter).WhereElementIsNotElementType()
                    .ToElementIds().Select(id => id.ToString()));
            }
            catch (Exception ex)
            {
                _viewFilterNote = "Не удалось отфильтровать активный вид. " + OpeningUserMessage.FromException(ex);
            }
        }

        private void PrepareRow(OpeningHistoryRow row, OpeningHistoryRow previous)
        {
            row.IsHistoryExpanded = previous?.IsHistoryExpanded ?? row.IsHistoryExpanded;
            if (row.IsConfirmed)
                _confirmationNotes.Remove(row.UniqueId);

            string note;
            if (_confirmationNotes.TryGetValue(row.UniqueId, out note))
                row.CacheNote = string.IsNullOrEmpty(row.CacheNote) ? note : row.CacheNote + Environment.NewLine + note;
        }

        private void SetRows(OpeningHistoryRow[] rows)
        {
            var previous = _cardsById.Values.ToDictionary(row => row.UniqueId);
            foreach (string id in _cardsById.Keys)
                _tracker.Remove(id);

            _cardsById.Clear();
            foreach (OpeningHistoryRow row in rows)
            {
                OpeningHistoryRow old;
                previous.TryGetValue(row.UniqueId, out old);
                PrepareRow(row, old);

                _cardsById.Add(row.ElementId, row);
                _tracker.Set(row);
            }

            // Полная замена только при полном сборе; ApplyRows сохраняет коллекцию и остальные карточки.
            _cards = new ObservableCollection<OpeningHistoryRow>(rows);
            Items = CollectionViewSource.GetDefaultView(_cards);
            Items.Filter = item => (!OnlySelected || _selectedIds.Contains(((OpeningHistoryRow)item).ElementId))
                && _filterSettings.Matches((OpeningHistoryRow)item, Filter, _visibleIds);

            UpdateRowsSnapshot();
        }

        private void ApplyRows(ElementId[] ids, OpeningHistoryRow[] rows)
        {
            var replacements = rows.ToDictionary(row => row.ElementId);
            foreach (ElementId id in ids)
            {
                string key = id.ToString();
                OpeningHistoryRow previous;
                OpeningHistoryRow current;
                _cardsById.TryGetValue(key, out previous);
                replacements.TryGetValue(key, out current);

                if (current == null)
                {
                    // Удалённое отверстие или семейство, больше не подходящее под фильтр.
                    if (previous != null)
                    {
                        _cards.Remove(previous);
                        _confirmationNotes.Remove(previous.UniqueId);
                    }

                    _cardsById.Remove(key);
                    _tracker.Remove(key);
                    continue;
                }

                PrepareRow(current, previous?.UniqueId == current.UniqueId ? previous : null);
                if (previous == null)
                {
                    // Порядок по числовому ID без преобразования в int (Revit 2024: 64 бита).
                    int index = 0;
                    while (index < _cards.Count && CompareIds(_cards[index].ElementId, key) < 0)
                        ++index;

                    _cards.Insert(index, current);
                }
                else
                    _cards[_cards.IndexOf(previous)] = current;

                _cardsById[key] = current;
                _tracker.Set(current);
            }

            UpdateRowsSnapshot();
        }

        private static int CompareIds(string left, string right)
        {
            int length = left.Length.CompareTo(right.Length);
            return length == 0 ? string.CompareOrdinal(left, right) : length;
        }

        private void UpdateRowsSnapshot()
        {
            _rows = _cards.ToArray();
            _elementIds = new HashSet<string>(_cardsById.Keys);
            UpdateCounts();
            CommandManager.InvalidateRequerySuggested();
        }


        private void UpdateCounts()
        {
            OnPropertyChanged(nameof(CountText));
            OnPropertyChanged(nameof(UnconfirmedCountText));
            OnPropertyChanged(nameof(HasUnconfirmed));
            OnPropertyChanged(nameof(SelectionFilterNote));
        }

        private void QueueConfirmAll()
        {
            if (!ConfirmAllCommand.CanExecute(null))
                return;

            Document document = _document;
            int generation = _generation;
            _confirmAllQueued = true;
            ActionStatus = "Подтверждение всех отверстий текущей модели…";
            CommandManager.InvalidateRequerySuggested();

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new HistoryCommand(document, app =>
            {
                if (generation != _generation)
                    return;

                try
                {
                    OpeningConfirmationResult result = OpeningConfirmationService.ConfirmAll(document);
                    _confirmationNotes = result.Errors;
                    // Отдельный итог не пропадёт при обновлении списка через DocumentChanged.
                    ActionStatus = result.Summary;
                    RequestElements(_rows.Select(row => row.RevitId));
                }
                catch (UnauthorizedAccessException)
                {
                    IsBimUser = false;
                    throw;
                }
                finally
                {
                    _confirmAllQueued = false;
                    CommandManager.InvalidateRequerySuggested();
                }
            }, message =>
            {
                if (generation != _generation)
                    return;

                _confirmAllQueued = false;
                ActionStatus = message;
                CommandManager.InvalidateRequerySuggested();
            }));
        }

        private enum ElementAction { Show, Select, Confirm }

        private void QueueElement(OpeningHistoryRow row, ElementAction action)
        {
            if (_suspended || row == null || _document == null || !_document.IsValidObject)
                return;


            Document document = _document;
            int generation = _generation;
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new HistoryCommand(document, app =>
            {
                if (generation != _generation)
                    return;


                Element element = document.GetElement(row.UniqueId);
                if (!OpeningHistoryService.IsOpening(element))
                    throw new InvalidOperationException("Отверстие удалено или больше не соответствует фильтру. Обновите список.");
                if (action == ElementAction.Confirm)
                {
                    using (Transaction tx = new Transaction(document, "KPLN: Подтвердить отверстие"))
                    {
                        tx.Start();
                        OpeningHistoryService.Confirm(element);
                        if (tx.Commit() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Revit отменил запись подтверждения");
                    }
                    _confirmationNotes.Remove(row.UniqueId);
                    RequestElements(new[] { element.Id });
                }
                else if (action == ElementAction.Show)
                {
                    OpeningZoomService.Show(app.ActiveUIDocument, element);
                    Status = "Показано отверстие ID " + row.ElementId;
                }
                else
                {
                    // Как выбор по ID в Clashes: только выделение, без изменения вида и масштаба.
                    app.ActiveUIDocument.Selection.SetElementIds(new[] { element.Id });
                    Status = "Выбрано отверстие ID " + row.ElementId;
                }
            }, message =>
            {
                if (generation == _generation)
                    Status = message;
            }));
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private sealed class DocumentCreations
        {
            internal Document Document { get; }
            internal Dictionary<string, OpeningCreationInfo> Entries { get; } = new Dictionary<string, OpeningCreationInfo>();

            internal DocumentCreations(Document document)
            {
                Document = document;
            }
        }


        private sealed class HistoryCommand : IExecutableCommand
        {
            private readonly Document _document;
            private readonly Action<UIApplication> _action;
            private readonly Action<string> _error;
            internal HistoryCommand(Document document, Action<UIApplication> action, Action<string> error)
            {
                _document = document;
                _action = action;
                _error = error;
            }


            public Result Execute(UIApplication app)
            {
                try
                {
                    Document activeDocument = app.ActiveUIDocument?.Document;
                    if (!_document.IsValidObject || activeDocument == null || !activeDocument.IsValidObject
                        || !_document.Equals(activeDocument))
                        throw new InvalidOperationException("Активная модель изменилась. Обновите список.");
                    _action(app);
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    _error(OpeningUserMessage.FromException(ex));
                    return Result.Cancelled;
                }
            }
        }
    }
}

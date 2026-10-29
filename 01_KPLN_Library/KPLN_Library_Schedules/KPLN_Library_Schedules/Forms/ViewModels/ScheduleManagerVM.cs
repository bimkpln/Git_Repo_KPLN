using Autodesk.Revit.DB;
using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Services.Schedules;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Library_Schedules.Forms.ViewModels
{
    public sealed class ScheduleManagerVM : ScheduleObservableObject
    {
        private readonly ScheduleParameterCatalog _parameterCatalog;
        private bool _availableParametersLoaded;
        private enum CopyMode { Edit, Manual, Parameter }
        private CopyMode _copyMode = CopyMode.Edit;
        private bool _groupByPrefix;
        private List<ScheduleCopyValue> _rawCopyValues;
        private string _prefixSeparator = ".";
        private string _prefixParts = "2";
        public bool GroupByPrefix { get => _groupByPrefix; set { Change(ref _groupByPrefix, value); PrefixSettingsChanged(); } }
        public string PrefixSeparator { get => _prefixSeparator; set { Change(ref _prefixSeparator, value); PrefixSettingsChanged(); } }
        public string PrefixParts { get => _prefixParts; set { Change(ref _prefixParts, value); PrefixSettingsChanged(); } }
        private void PrefixSettingsChanged()
        {
            CopyValues.Clear();
            ValuesStatus = "Настройки изменены. Нажмите «Получить значения».";
            SyncCopyRule();
            Invalidate();
            if (_rawCopyValues != null)
            {
                try { RebuildCopyValues(); }
                catch (InvalidOperationException error) { ValuesStatus = error.Message; }
            }
        }
        private ScheduleParameterChoice _copyParameter;
        private string _selectionKey;
        private string _valuesStatus = "Выберите параметр и получите его значения.";
        private readonly List<BatchScheduleFilter> _replacedRules = new List<BatchScheduleFilter>();
        public ObservableCollection<ScheduleCopyValue> CopyValues { get; } = new ObservableCollection<ScheduleCopyValue>();
        public ObservableCollection<ScheduleParameterChoice> CopyParameters { get; } = new ObservableCollection<ScheduleParameterChoice>();
        public bool CanCopy => Schedules.Count(item => item.Include) == 1;
        public bool IsParameterCopy => CreateCopies && ByParameter && CanCopy;
        public string ValuesStatus { get => _valuesStatus; private set => Change(ref _valuesStatus, value); }
        public ScheduleParameterChoice CopyParameter
        {
            get => _copyParameter;
            set
            {
                if (ReferenceEquals(_copyParameter, value)) return;
                Change(ref _copyParameter, value);
                _rawCopyValues = null;
                CopyValues.Clear();
                ValuesStatus = "Нажмите «Получить значения». Пустые значения пропускаются.";
                SyncCopyRule();
                Invalidate();
            }
        }
        private bool _ready;
        private bool _selectingAll;
        private string _namePattern = "{Исходное}_{Значение}";
        private string _status = "Выберите спецификации. Настройте копирование и фильтры, затем проверьте результат.";
        public Document Document { get; }
        public bool IsByLevels { get; private set; }
        public event Action RulesChanged;
        public ObservableCollection<ScheduleItem> Schedules { get; } = new ObservableCollection<ScheduleItem>();
        public ObservableCollection<ScheduleItem> VisibleSchedules { get; } = new ObservableCollection<ScheduleItem>();
        private string _scheduleSearch = "";
        public string ScheduleSearch
        {
            get => _scheduleSearch;
            set
            {
                if (_scheduleSearch == value) return;
                Change(ref _scheduleSearch, value);
                RefreshVisibleSchedules();
            }
        }

        private void RefreshVisibleSchedules()
        {
            string query = (ScheduleSearch ?? "").Trim();
            VisibleSchedules.Clear();
            foreach (var item in Schedules.Where(item => query.Length == 0
                || item.Name.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0))
                VisibleSchedules.Add(item);
        }

        public ObservableCollection<ScheduleParameterChoice> Parameters { get; } = new ObservableCollection<ScheduleParameterChoice>();
        public ObservableCollection<BatchScheduleFilter> BatchRules { get; } = new ObservableCollection<BatchScheduleFilter>();
        public ObservableCollection<BatchScheduleSort> SortRules { get; } = new ObservableCollection<BatchScheduleSort>();
        public ObservableCollection<ScheduleParameterChoice> SortParameters { get; } = new ObservableCollection<ScheduleParameterChoice>();
        private bool _sortParametersLoaded;
        public ObservableCollection<string> SortPositions { get; } = new ObservableCollection<string>();
        private readonly Dictionary<BatchScheduleSort, string> _sortPositions = new Dictionary<BatchScheduleSort, string>();
        private bool _updatingSortPositions;

        public ObservableCollection<ScheduleOperation> Operations { get; } = new ObservableCollection<ScheduleOperation>();
        public bool CreateCopies => _copyMode != CopyMode.Edit;
        public bool ManualCopy
        {
            get => _copyMode == CopyMode.Manual;
            set => SetCopyMode(value ? CopyMode.Manual : ManualCopy ? CopyMode.Edit : _copyMode);
        }
        public bool ByParameter
        {
            get => _copyMode == CopyMode.Parameter;
            set => SetCopyMode(value ? CopyMode.Parameter : ByParameter ? CopyMode.Edit : _copyMode);
        }

        private void SetCopyMode(CopyMode mode)
        {
            if (!CanCopy) mode = CopyMode.Edit;
            if (_copyMode == mode) return;
            _copyMode = mode;

            if (mode == CopyMode.Manual && NamePattern == "{Исходное}_{Значение}")
                NamePattern = "{Исходное} — копия";
            if (mode == CopyMode.Parameter && NamePattern == "{Исходное} — копия")
                NamePattern = "{Исходное}_{Значение}";

            Notify(nameof(ManualCopy));
            Notify(nameof(ByParameter));
            Notify(nameof(CreateCopies));
            Notify(nameof(IsParameterCopy));
            SyncCopyRule();
            Invalidate();
        }
        public string NamePattern { get => _namePattern; set { Change(ref _namePattern, value); Invalidate(); } }
        public bool Ready { get => _ready; private set => Change(ref _ready, value); }
        public string Status { get => _status; private set => Change(ref _status, value); }
        public ScheduleExecutionReport ExecutionReport { get; private set; }
        public bool HasExecutionReport => ExecutionReport != null;

        public ScheduleManagerVM(Document document, ISet<ElementId> selected, IEnumerable<ViewSchedule> scope = null)
        {
            Document = document;
            _parameterCatalog = new ScheduleParameterCatalog(document);
            foreach (ViewSchedule schedule in (scope ?? new FilteredElementCollector(document).OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()).Where(s => !s.IsTemplate && !s.IsTitleblockRevisionSchedule).OrderBy(s => s.Name))
            {
                var item = new ScheduleItem { Schedule = schedule, Include = selected.Contains(schedule.Id) };
                ScheduleService.ReadFilters(document, item);
                item.PropertyChanged += (s, e) =>
                {
                    Invalidate();
                    if (!_selectingAll && e.PropertyName == nameof(ScheduleItem.Include)) RefreshParameters();
                };
                Schedules.Add(item);
            }
            RefreshVisibleSchedules();
            RefreshParameters();
        }

        public void ConfigureForLevels(IEnumerable<LevelScheduleFilter> filters, IEnumerable<BatchScheduleSort> sorts)
        {
            IsByLevels = true;
            BatchRules.Clear();
            foreach (var rule in filters)
            {
                var parameter = Parameters.FirstOrDefault(item => item.Key == rule.Parameter?.Key);
                if (parameter == null && rule.Parameter != null) Parameters.Add(rule.Parameter);
                else if (parameter != null) rule.Parameter = parameter;
                rule.PropertyChanged += (sender, args) => Invalidate();
                BatchRules.Add(rule);
            }
            if (sorts != null)
            {
                SortRules.Clear();
                _sortPositions.Clear();
                foreach (var rule in sorts)
                {
                    var parameter = SortParameters.FirstOrDefault(item => item.Key == rule.Parameter?.Key);
                    if (parameter == null && rule.Parameter != null) SortParameters.Add(rule.Parameter);
                    else if (parameter != null) rule.Parameter = parameter;
                    AttachSort(rule);
                }
                UpdateSortPositions();
            }
        }

        public void Invalidate()
        {
            Ready = false;
            RulesChanged?.Invoke();
            Status = "После изменений нажмите «Проверить и показать результат».";
        }

        public void RefreshParameters()
        {
            RestoreCopyRules();
            var sources = Schedules.Where(s => s.Include).ToList();
            string selection = string.Join("|", sources.Select(item => item.Schedule.UniqueId).OrderBy(key => key));
            if (selection != _selectionKey)
            {
                _selectionKey = selection;
                _rawCopyValues = null;
                _copyParameter = null;
                Notify(nameof(CopyParameter));
                CopyValues.Clear();
                ValuesStatus = "Выберите параметр и получите его значения.";
            }
            if (!CanCopy) SetCopyMode(CopyMode.Edit);
            Notify(nameof(CanCopy));
            Notify(nameof(IsParameterCopy));
            foreach (var existing in BatchRules.Where(r => r.IsExisting).ToList()) BatchRules.Remove(existing);
            var selectedKeys = BatchRules.ToDictionary(rule => rule, rule => rule.Parameter?.Key);
            Parameters.Clear();
            CopyParameters.Clear();
            _availableParametersLoaded = BatchRules.Any(rule => !rule.IsExisting);
            foreach (ScheduleParameterChoice parameter in ScheduleBatchService.CollectParameters(Document, sources, _parameterCatalog, _availableParametersLoaded))
            {
                Parameters.Add(parameter);
                if (ScheduleCopyService.Supports(parameter)) CopyParameters.Add(parameter);
            }
            foreach (BatchScheduleFilter rule in BatchRules)
            {
                string key = selectedKeys[rule];
                rule.Parameter = Parameters.FirstOrDefault(p => p.Key == key);
            }
            LoadCommonFilters(sources);
            SortRules.Clear();
            SortParameters.Clear();
            _sortParametersLoaded = false;
            _sortPositions.Clear();
            foreach (var row in ScheduleSortingService.ReadCommon(Document, sources))
            {
                var parameter = SortParameters.FirstOrDefault(item => item.Key == row.Parameter.Key);
                if (parameter == null) SortParameters.Add(row.Parameter);
                else row.Parameter = parameter;
                AttachSort(row);
            }
            UpdateSortPositions();
            SyncCopyRule();
        }

        private void AttachSort(BatchScheduleSort row)
        {
            _sortPositions[row] = row.Position;
            row.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == nameof(BatchScheduleSort.Position) && !_updatingSortPositions)
                {
                    _updatingSortPositions = true;
                    try
                    {
                        string old = _sortPositions[row];
                        if (row.Position != "Исходное")
                            foreach (var other in SortRules.Where(item => item != row && item.Enabled && !item.Deleted && item.Position == row.Position))
                            {
                                other.Position = old;
                                _sortPositions[other] = old;
                            }
                        _sortPositions[row] = row.Position;
                    }
                    finally { _updatingSortPositions = false; }
                }
                if (args.PropertyName == nameof(BatchScheduleSort.Enabled) || args.PropertyName == nameof(BatchScheduleSort.Deleted)) UpdateSortPositions();
                Invalidate();
            };
            SortRules.Add(row);
        }

        private void UpdateSortPositions()
        {
            var sources = Schedules.Where(item => item.Include).ToList();
            int count = sources.Count == 0 ? 0 : sources.Min(item => item.Schedule.Definition.GetSortGroupFields().Count)
                + SortRules.Count(row => !row.IsExisting && row.Enabled)
                - SortRules.Count(row => row.IsExisting && row.Enabled && row.Deleted);
            var desired = new[] { "Исходное" }.Concat(Enumerable.Range(1, Math.Max(0, count)).Select(index => index.ToString())).ToList();
            if (SortPositions.SequenceEqual(desired)) return;
            // Перезаполнение ItemsSource сбрасывает SelectedItem ComboBox: сохраняем значения строк.
            var positions = SortRules.ToDictionary(row => row, row => row.Position);
            _updatingSortPositions = true;
            try
            {
                SortPositions.Clear();
                foreach (string position in desired) SortPositions.Add(position);
                foreach (var pair in positions)
                {
                    pair.Key.Position = desired.Contains(pair.Value) ? pair.Value : "Исходное";
                    _sortPositions[pair.Key] = pair.Key.Position;
                }
            }
            finally { _updatingSortPositions = false; }
        }

        public void EnsureSortParameters()
        {
            if (_sortParametersLoaded) return;
            var sources = Schedules.Where(item => item.Include).ToList();
            if (sources.Count == 0) throw new InvalidOperationException("Сначала выберите спецификации.");
            var known = new HashSet<string>(SortParameters.Select(parameter => parameter.Key));
            foreach (var parameter in ScheduleSortingService.CollectParameters(Document, sources, _parameterCatalog))
            {
                if (!known.Add(parameter.Key)) continue;
                int position = 0;
                while (position < SortParameters.Count && string.Compare(SortParameters[position].Name, parameter.Name, StringComparison.CurrentCulture) <= 0)
                    position++;
                SortParameters.Insert(position, parameter);
            }
            _sortParametersLoaded = true;
        }

        public void AddSort()
        {
            var sources = Schedules.Where(item => item.Include).ToList();
            if (sources.Count == 0) throw new InvalidOperationException("Сначала выберите спецификации.");
            int added = SortRules.Count(row => !row.IsExisting && row.Enabled);
            int removed = SortRules.Count(row => row.IsExisting && row.Enabled && row.Deleted);
            if (sources.Any(item => item.Schedule.Definition.GetSortGroupFields().Count + added - removed >= 4))
                throw new InvalidOperationException("Revit допускает не более четырёх параметров сортировки. Сначала удалите ненужную сортировку.");
            EnsureSortParameters();
            var first = SortParameters.FirstOrDefault(parameter => !SortRules.Any(row => !row.Deleted && row.Parameter?.Key == parameter.Key));
            if (first == null) throw new InvalidOperationException("Нет других общих параметров для сортировки.");
            var rule = new BatchScheduleSort { Parameter = first };
            AttachSort(rule);
            UpdateSortPositions();
            rule.Position = SortPositions.Last();
            Invalidate();
        }

        public void RemoveSort(BatchScheduleSort row)
        {
            if (row == null) return;
            if (row.IsExisting) { row.Deleted = !row.Deleted; row.Enabled = true; }
            else { SortRules.Remove(row); _sortPositions.Remove(row); }
            UpdateSortPositions();
            Invalidate();
        }

        private void LoadCommonFilters(List<ScheduleItem> sources)
        {
            foreach (var rule in ScheduleBatchService.ReadCommonFilters(Document, sources, Parameters))
            {
                // Существующее общее поле должно быть видно и до раскрытия полного каталога.
                if (!Parameters.Any(parameter => parameter.Key == rule.Parameter.Key)) Parameters.Add(rule.Parameter);
                rule.PropertyChanged += (sender, args) => Invalidate();
                BatchRules.Add(rule);
            }
        }

        public void RemoveRule(BatchScheduleFilter rule)
        {
            if (rule == null) return;
            if (rule.IsAutomatic) { Status = "Автоматический фильтр управляется параметром копирования на первой вкладке."; return; }
            if (rule.IsExisting)
            {
                rule.Deleted = !rule.Deleted;
                rule.Enabled = true;
            }
            else BatchRules.Remove(rule);
            Invalidate();
        }

        public void SelectSchedules(bool selected)
        {
            _selectingAll = true;
            try { foreach (ScheduleItem item in VisibleSchedules.ToList()) item.Include = selected; }
            finally { _selectingAll = false; }
            RefreshParameters();
            Invalidate();
        }

        public void EnsureAvailableParameters()
        {
            if (!_availableParametersLoaded)
            {
                // Сохраняем объекты существующих строк и их правки при заполнении списка.
                var known = new HashSet<string>(Parameters.Select(parameter => parameter.Key));
                var sources = Schedules.Where(item => item.Include).ToList();
                foreach (var parameter in ScheduleBatchService.CollectParameters(Document, sources, _parameterCatalog))
                {
                    if (!known.Add(parameter.Key)) continue;
                    int position = 0;
                    while (position < Parameters.Count && string.Compare(Parameters[position].Name, parameter.Name, StringComparison.CurrentCulture) <= 0)
                        position++;
                    Parameters.Insert(position, parameter);
                    if (ScheduleCopyService.Supports(parameter))
                    {
                        int copyPosition = Parameters.Take(position).Count(ScheduleCopyService.Supports);
                        CopyParameters.Insert(copyPosition, parameter);
                    }
                }
                _availableParametersLoaded = true;
            }
        }

        public void AddRule()
        {
            EnsureAvailableParameters();
            BatchScheduleFilter rule = IsByLevels
                ? new LevelScheduleFilter { Parameter = Parameters.FirstOrDefault(), Value = "{Этаж}" }
                : new BatchScheduleFilter { Parameter = Parameters.FirstOrDefault() };
            rule.PropertyChanged += (s, e) => Invalidate();
            BatchRules.Add(rule);
            Invalidate();
        }

        private void RestoreCopyRules()
        {
            foreach (var automatic in BatchRules.Where(rule => rule.IsAutomatic).ToList()) BatchRules.Remove(automatic);
            foreach (var rule in _replacedRules) BatchRules.Add(rule);
            _replacedRules.Clear();
        }

        private void SyncCopyRule()
        {
            RestoreCopyRules();
            if (!IsParameterCopy || CopyParameter == null) return;
            foreach (var rule in BatchRules.Where(rule => rule.Parameter?.Key == CopyParameter.Key).ToList())
            {
                _replacedRules.Add(rule);
                BatchRules.Remove(rule);
            }
            BatchRules.Add(new BatchScheduleFilter
            {
                IsAutomatic = true, IsExisting = _replacedRules.Any(rule => rule.IsExisting),
                Parameter = CopyParameter, Condition = GroupByPrefix ? "Начинается с" : "Равно",
                Value = "{Значение}",
                Kind = CopyParameter.SuggestedKind,
                Visibility = _replacedRules.Select(rule => rule.Visibility).Distinct().Count() == 1 ? _replacedRules[0].Visibility : "Исходное"
            });
            BatchRules.Last().PropertyChanged += (sender, args) => Invalidate();
        }

        public void ReadCopyValues()
        {
            if (!IsParameterCopy) throw new InvalidOperationException("Для копирования выберите ровно одну спецификацию.");
            if (CopyParameter == null) throw new InvalidOperationException("Выберите параметр копирования.");
            Invalidate();
            CopyValues.Clear();
            ValuesStatus = "Чтение значений…";
            _rawCopyValues = null;
            var source = Schedules.Single(item => item.Include);
            try
            {
                int parts = 0;
                if (GroupByPrefix && (CopyParameter.SuggestedKind != "Текст"
                    || !int.TryParse(PrefixParts, out parts) || parts < 1 || parts > 20 || string.IsNullOrEmpty(PrefixSeparator)))
                    throw new InvalidOperationException("Для группировки выберите текстовый параметр, разделитель и число частей от 1 до 20.");
                _rawCopyValues = ScheduleCopyService.ReadValues(Document, source.Schedule, CopyParameter);
                RebuildCopyValues();
            }
            catch { ValuesStatus = "Значения не получены. Исправьте причину и повторите чтение."; throw; }
            finally { ScheduleService.ReadFilters(Document, source); }
        }

        private void RebuildCopyValues()
        {
            CopyValues.Clear();
            int parts;
            if (GroupByPrefix && (CopyParameter?.SuggestedKind != "Текст" || !int.TryParse(PrefixParts, out parts)))
                throw new InvalidOperationException("Для группировки выберите текстовый параметр и целое число частей от 1 до 20.");
            int.TryParse(PrefixParts, out parts);
            var values = GroupByPrefix
                ? ScheduleCopyService.GroupPrefixes(_rawCopyValues, PrefixSeparator, parts)
                : _rawCopyValues;
            foreach (var value in values)
            {
                var row = new ScheduleCopyValue { Storage = value.Storage, Text = value.Text,
                    Integer = value.Integer, Number = value.Number, ElementIdValue = value.ElementIdValue,
                    Display = value.Display, ElementCount = value.ElementCount, IsPrefix = value.IsPrefix,
                    Selected = false };
                row.PropertyChanged += (sender, args) => Invalidate();
                CopyValues.Add(row);
            }
            ValuesStatus = "Значений: " + CopyValues.Count
                + ". Список без фильтров исходной спецификации. Отметьте нужные значения. Условия копий настройте на вкладке 2.";
        }

        public void SelectValues(bool selected)
        {
            foreach (var row in CopyValues) row.Selected = selected;
            Invalidate();
        }

        public void BuildPreview()
        {
            Ready = false;
            Operations.Clear();
            try
            {
                var sources = Schedules.Where(item => item.Include).ToList();
                if (sources.Count == 0) throw new InvalidOperationException("Выберите хотя бы одну спецификацию.");
                if (CreateCopies && sources.Count != 1) throw new InvalidOperationException("Копировать можно только одну исходную спецификацию.");
                var values = IsParameterCopy ? CopyValues.Where(value => value.Selected).ToList() : new List<ScheduleCopyValue> { null };
                if (IsParameterCopy && (CopyParameter == null || values.Count == 0))
                    throw new InvalidOperationException("Выберите параметр, получите значения и отметьте хотя бы одно.");
                if (BatchRules.Any(rule => rule.Enabled && rule.Parameter == null))
                    throw new InvalidOperationException("Выберите параметр для каждого включённого правила.");
                if (IsParameterCopy && BatchRules.Any(rule => !rule.IsAutomatic && rule.Enabled && rule.Parameter?.Key == CopyParameter.Key))
                    throw new InvalidOperationException("Параметром копирования управляет автоматическая строка. Удалите дополнительное правило этого параметра.");
                if (SortRules.Any(rule => rule.Enabled && rule.Parameter == null))
                    throw new InvalidOperationException("Выберите параметр для каждой включённой сортировки.");
                if (!CreateCopies && !BatchRules.Any(rule => rule.HasChanges) && !SortRules.Any(rule => rule.HasChanges))
                    throw new InvalidOperationException("Не заданы изменения фильтров, сортировки или видимости.");
                if (CreateCopies) ScheduleMaskCatalog.Constructor(true, IsParameterCopy).Validate(NamePattern);
                foreach (var rule in BatchRules.Where(rule => rule.HasChanges && !rule.IsAutomatic))
                    ScheduleMaskCatalog.Constructor(allowOriginal: rule.IsExisting).Validate(rule.Value);

                foreach (var source in sources)
                {
                    int index = 0;
                    foreach (var value in values)
                        Operations.Add(new ScheduleOperation
                        {
                            Source = source, Level = new ScheduleLevelRow(), Index = ++index,
                            CreateCopy = CreateCopies, CopyParameter = value == null ? null : CopyParameter,
                            CopyValue = value, TargetName = source.Name
                        });
                }
                var names = CreateCopies ? new HashSet<string>(new FilteredElementCollector(Document).OfClass(typeof(View)).Cast<View>().Select(view => view.Name), StringComparer.OrdinalIgnoreCase) : null;
                var entries = Operations.ToDictionary(operation => operation, operation =>
                    ScheduleBatchService.FilterEntriesForName(Document, operation.Source, BatchRules, operation.Level, operation.Index, operation));
                var common = new HashSet<Tuple<string, string, string>>(entries.Values.First());
                if (Operations.Count > 1)
                    foreach (var entry in entries.Values.Skip(1)) common.IntersectWith(entry);
                else if (!IsParameterCopy) common.Clear();
                foreach (var operation in Operations.Where(item => item.CreateCopy))
                {
                    var individual = entries[operation].Where(entry => !common.Contains(entry)
                        || entry.Item1 == operation.CopyParameter?.Key);
                    string filters = string.Join("/", individual.Select(entry => entry.Item3).Distinct());
                    string name = ScheduleNaming.Expand(NamePattern, null, null, null, null, operation.SourceName,
                        operation.Index, filters, operation.CopyParameter?.Name, operation.CopyValue?.Display);
                    string error = ScheduleSupport.InvalidName(name);
                    if (error != null) throw new InvalidOperationException(error);
                    string prefix = name;
                    int suffix = 2;
                    while (!names.Add(name)) name = prefix + " (" + suffix++ + ")";
                    operation.TargetName = name;
                }
                Ready = ScheduleBatchService.Preview(Document, Operations.ToList(), BatchRules.ToList(), SortRules.ToList());
                Status = Ready ? "Проверено операций: " + Operations.Count + ". Изменения ещё не применены. Проверьте имена и значения, затем нажмите «Выполнить»."
                    : "Есть ошибки. Исправьте настройки и повторите проверку.";
            }
            catch (Exception ex) { Status = ex.Message; }
        }

        public void Execute()
        {
            if (!Ready) return;
            Ready = false;
            var timer = System.Diagnostics.Stopwatch.StartNew();
            string error = null;
            bool completed = false;
            try
            {
                ScheduleBatchService.Execute(Document, Operations.ToList(), BatchRules.ToList(), SortRules.ToList());
                completed = true;
                foreach (ScheduleItem item in Schedules) ScheduleService.ReadFilters(Document, item);
                _replacedRules.Clear();
                BatchRules.Clear();
                CopyValues.Clear();
                ValuesStatus = "После выполнения обновите значения перед новым копированием.";
                RefreshParameters();
            }
            catch (Exception ex)
            {
                error = (completed ? "Не удалось обновить окно: " : "Ошибка выполнения: ") + ex.Message;
            }
            finally
            {
                timer.Stop();
                ExecutionReport = ScheduleExecutionReport.Create(Operations, timer.Elapsed, error);
                Notify(nameof(ExecutionReport));
                Notify(nameof(HasExecutionReport));
                Status = ExecutionReport.Summary + (error == null ? "" : " " + error);
            }
        }
    }
}


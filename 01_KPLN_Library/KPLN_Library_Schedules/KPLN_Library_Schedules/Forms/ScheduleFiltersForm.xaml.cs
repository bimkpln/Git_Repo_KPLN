using Autodesk.Revit.DB;
using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Services.Schedules;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using KPLN_Library_Schedules.Forms.ViewModels;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace KPLN_Library_Schedules.Forms
{
    public partial class ScheduleFiltersForm : Window
    {
        private readonly ScheduleItem _source;
        private readonly Document _document;
        private readonly List<ScheduleLevelRow> _levels;
        private bool _refreshPending;
        private bool _closed;
        public string ScheduleName => _source.Name;
        public ScheduleManagerVM Editor { get; }
        public IEnumerable<LevelScheduleFilter> Rules => Editor.BatchRules.Cast<LevelScheduleFilter>();
        public ObservableCollection<LevelFilterValueRow> Values { get; } = new ObservableCollection<LevelFilterValueRow>();

        public ScheduleFiltersForm(ScheduleItem source, Document document, IEnumerable<ScheduleLevelRow> levels)
        {
            _source = source;
            _document = document;
            _levels = levels.ToList();
            Editor = new ScheduleManagerVM(document, new HashSet<ElementId> { source.Schedule.Id }, new[] { source.Schedule });
            Editor.ConfigureForLevels(ScheduleLevelFilterService.ReadRules(document, source),
                source.LevelSorts?.Select(ScheduleSortingService.Clone));
            InitializeComponent();
            DataContext = this;
            Editor.RulesChanged += QueueRefresh;
            Closed += (sender, args) => { _closed = true; Editor.RulesChanged -= QueueRefresh; };
            BuildValues();
        }
        private void QueueRefresh()
        {
            if (_refreshPending || _closed) return;
            _refreshPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _refreshPending = false;
                if (!_closed) BuildValues();
            }), DispatcherPriority.Background);
        }

        private void BuildValues()
        {
            Values.Clear();
            int index = 0;
            bool errors = false;
            foreach (var level in _levels)
            {
                index++;
                foreach (var rule in Rules)
                {
                    bool keep = rule.IsExisting && (!rule.Enabled || !rule.HasChanges);
                    bool active = rule.Enabled && !rule.Deleted;
                    var row = new LevelFilterValueRow
                    {
                        Rule = rule, Level = level,
                        Field = (rule.Parameter?.Name ?? "Выберите параметр") + " / "
                            + (keep ? rule.OriginalCondition : rule.Condition),
                        Mode = !rule.Enabled ? (rule.IsExisting ? "Без изменений" : "Отключён")
                            : rule.Deleted ? "Будет удалён" : rule.Mode,
                        CanEdit = active && rule.NeedsValue
                            && (rule.Mode == FilterModes.Pattern || rule.Mode == FilterModes.Manual)
                    };
                    try { row.Value = ScheduleLevelFilterService.ResolveText(rule, level, index); }
                    catch (Exception ex)
                    {
                        row.Value = "Ошибка: " + ex.Message;
                        row.CanEdit = false;
                        errors = true;
                    }
                    row.PropertyChanged += (sender, args) =>
                    {
                        if (args.PropertyName == nameof(LevelFilterValueRow.Value) && row.CanEdit)
                            row.Rule.Overrides[row.Level.Key] = row.Value ?? "";
                    };
                    Values.Add(row);
                }
            }
            PreviewStatus.Text = _levels.Count == 0 ? "Выберите уровни в основном окне, чтобы увидеть значения по этажам."
                : errors ? "Есть ошибки в масках. Исправьте правила; таблица обновится автоматически."
                : "Предпросмотр по " + _levels.Count + " ур. Правки ещё не применены к модели.";
        }

        private void Commit()
        {
            RulesEditor.Commit();
            foreach (var grid in new[] { ValuesGrid })
            {
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true))
                    throw new InvalidOperationException("Исправьте значение в редактируемой ячейке.");
            }
        }

        private void BeginValueEdit(object sender, DataGridBeginningEditEventArgs args)
        {
            var row = args.Row.Item as LevelFilterValueRow;
            if (row == null || !row.CanEdit || _refreshPending)
                args.Cancel = true;
        }

        private void RebuildValues(object sender, RoutedEventArgs args)
        {
            RunBusy(() =>
            {
                Commit();
                foreach (var row in Values.Where(value => value.CanEdit && value.Rule.Mode == FilterModes.Manual))
                    row.Rule.Overrides[row.Level.Key] = row.Value ?? "";
                foreach (var rule in Rules.Where(rule => rule.Mode == FilterModes.Pattern))
                    rule.Overrides.Clear();
                BuildValues();
            });
        }

        private void Apply(object sender, RoutedEventArgs args)
        {
            RunBusy(() =>
            {
                Commit();
                if (_levels.Count == 0)
                    throw new InvalidOperationException("Выберите хотя бы один уровень в основном окне.");

                // Ручные значения сохраняются, включая пустую строку.
                if (_refreshPending)
                    BuildValues();
                foreach (var row in Values.Where(value => value.CanEdit && value.Rule.Mode == FilterModes.Manual))
                    row.Rule.Overrides[row.Level.Key] = row.Value ?? "";

                ScheduleLevelFilterService.Validate(_document, _source, Rules.ToList(),
                    _levels.Select((level, index) => new ScheduleOperation { Level = level, Index = index + 1 }), Editor.SortRules.ToList());
                _source.LevelFilters = Rules.Select(ScheduleLevelFilterService.Clone).ToList();
                _source.LevelSorts = Editor.SortRules.Select(ScheduleSortingService.Clone).ToList();
                DialogResult = true;
            });
        }

        private void RunBusy(Action action)
        {
            var previous = System.Windows.Input.Mouse.OverrideCursor;
            try
            {
                System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                action();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Title); }
            finally { System.Windows.Input.Mouse.OverrideCursor = previous; }
        }
    }
}

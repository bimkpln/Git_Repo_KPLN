using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using KPLN_ViewsAndLists_Ribbon.Services.Schedules;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ScheduleFiltersForm : Window
    {
        private readonly ScheduleItem _source;
        private readonly Document _document;
        private readonly List<ScheduleLevelRow> _levels;
        private bool _parametersLoaded;
        private bool _refreshPending;
        private bool _closed;
        public string ScheduleName => _source.Name;
        public ObservableCollection<LevelScheduleFilter> Rules { get; } = new ObservableCollection<LevelScheduleFilter>();
        public ObservableCollection<ScheduleParameterChoice> Parameters { get; } = new ObservableCollection<ScheduleParameterChoice>();
        public ObservableCollection<LevelFilterValueRow> Values { get; } = new ObservableCollection<LevelFilterValueRow>();

        public ScheduleFiltersForm(ScheduleItem source, Document document, IEnumerable<ScheduleLevelRow> levels)
        {
            _source = source;
            _document = document;
            _levels = levels.ToList();
            foreach (var rule in ScheduleLevelFilterService.ReadRules(document, source))
            {
                rule.PropertyChanged += RuleChanged;
                Rules.Add(rule);
            }

            InitializeComponent();
            DataContext = this;
            Closed += (sender, args) =>
            {
                _closed = true;
                foreach (var rule in Rules)
                    rule.PropertyChanged -= RuleChanged;
            };
            if (Rules.Any(rule => !rule.IsExisting))
                RunBusy(EnsureParameters);
            BuildValues();
        }

        private void EnsureParameters()
        {
            if (_parametersLoaded)
                return;

            // Каталог и поиск параметров совпадают с конструктором спецификаций.
            var source = new ScheduleItem { Schedule = _source.Schedule };
            foreach (var parameter in ScheduleBatchService.CollectParameters(_document, new[] { source }))
                Parameters.Add(parameter);
            foreach (var rule in Rules.Where(rule => !rule.IsExisting))
            {
                var parameter = Parameters.FirstOrDefault(item => item.Key == rule.Parameter?.Key);
                if (parameter != null)
                {
                    rule.PropertyChanged -= RuleChanged;
                    rule.Parameter = parameter;
                    rule.PropertyChanged += RuleChanged;
                }
            }
            _parametersLoaded = true;
        }

        private void AddRule(object sender, RoutedEventArgs args)
        {
            RunBusy(() =>
            {
                Commit();
                EnsureParameters();
                var rule = new LevelScheduleFilter { Parameter = Parameters.FirstOrDefault(), Value = "{Этаж}" };
                rule.PropertyChanged += RuleChanged;
                Rules.Add(rule);
                RulesGrid.SelectedItem = rule;
                RulesGrid.ScrollIntoView(rule);
                BuildValues();
            });
        }

        private void RemoveRule(object sender, RoutedEventArgs args)
        {
            RunBusy(() =>
            {
                Commit();
                var rule = RulesGrid.SelectedItem as LevelScheduleFilter;
                if (rule == null)
                    return;

                if (rule.IsExisting)
                {
                    rule.Deleted = !rule.Deleted;
                    rule.Enabled = true;
                }
                else
                {
                    rule.PropertyChanged -= RuleChanged;
                    Rules.Remove(rule);
                }
                BuildValues();
            });
        }

        private void EditMask(object sender, RoutedEventArgs args)
        {
            var rule = ((FrameworkElement)sender).DataContext as LevelScheduleFilter;
            if (rule == null)
                return;

            var dialog = new ScheduleMaskForm(rule.Value,
                ScheduleMaskCatalog.CopyByLevels(allowOriginal: rule.IsExisting)) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                rule.Value = dialog.Pattern;
                rule.Mode = FilterModes.Pattern;
            }
        }

        private void RuleChanged(object sender, PropertyChangedEventArgs args)
        {
            if (args.PropertyName != nameof(LevelScheduleFilter.Parameter)
                && args.PropertyName != nameof(LevelScheduleFilter.Condition)
                && args.PropertyName != nameof(LevelScheduleFilter.Kind)
                && args.PropertyName != nameof(LevelScheduleFilter.Value)
                && args.PropertyName != nameof(LevelScheduleFilter.Mode)
                && args.PropertyName != nameof(LevelScheduleFilter.Enabled)
                && args.PropertyName != nameof(LevelScheduleFilter.Deleted))
                return;

            var rule = (LevelScheduleFilter)sender;
            if (args.PropertyName == nameof(LevelScheduleFilter.Value) || args.PropertyName == nameof(LevelScheduleFilter.Parameter))
                rule.Overrides.Clear();
            if (_refreshPending || _closed)
                return;

            _refreshPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _refreshPending = false;
                if (!_closed)
                    BuildValues();
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
            foreach (var grid in new[] { RulesGrid, ValuesGrid })
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
                    _levels.Select((level, index) => new ScheduleOperation { Level = level, Index = index + 1 }));
                _source.LevelFilters = Rules.Select(ScheduleLevelFilterService.Clone).ToList();
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

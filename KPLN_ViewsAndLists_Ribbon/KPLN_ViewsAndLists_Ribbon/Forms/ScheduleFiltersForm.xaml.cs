using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using FilterRule = KPLN_ViewsAndLists_Ribbon.Common.Schedules.FilterRule;
using KPLN_ViewsAndLists_Ribbon.Services.Schedules;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ScheduleFiltersForm : Window
    {
        private readonly ScheduleItem _source;
        private readonly Document _document;
        private readonly List<ScheduleLevelRow> _levels;
        private string _builtRules;
        public string ScheduleName => _source.Name;
        public ObservableCollection<FilterRule> Rules { get; } = new ObservableCollection<FilterRule>();
        public ObservableCollection<FilterValueRow> Values { get; } = new ObservableCollection<FilterValueRow>();

        public ScheduleFiltersForm(ScheduleItem source, Document document, IEnumerable<ScheduleLevelRow> levels)
        {
            _source = source;
            _document = document;
            _levels = levels.ToList();
            foreach (FilterRule rule in source.Filters)
            {
                var copy = new FilterRule
                {
                    Index = rule.Index, Original = rule.Original, FieldIdentity = rule.FieldIdentity,
                    FieldName = rule.FieldName, Value = rule.Value, ValueType = rule.ValueType, Mode = rule.Mode, Pattern = rule.Pattern
                };
                foreach (var pair in rule.Overrides) copy.Overrides.Add(pair.Key, pair.Value);
                Rules.Add(copy);
            }
            InitializeComponent();
            DataContext = this;
            try { BuildValues(false); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message + " Исправьте маску и пересчитайте значения.", "KPLN: фильтры"); }
        }

        private void EditMask(object sender, RoutedEventArgs args)
        {
            var rule = ((FrameworkElement)sender).DataContext as FilterRule;
            if (rule == null) return;

            var dialog = new ScheduleMaskForm(rule.Pattern) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                rule.Pattern = dialog.Pattern;
                rule.Mode = FilterModes.Pattern;
            }
        }

        private string Fingerprint() => string.Join("|", Rules.Select(r => r.Index + ":" + r.Mode + ":" + r.Pattern));

        private void BuildValues(bool resetMasks)
        {
            foreach (FilterValueRow value in Values.Where(v => v.Mode == FilterModes.Manual && v.Rule.Mode == FilterModes.Manual))
                value.Rule.Overrides[value.Level.Key] = value.Value ?? "";
            if (resetMasks)
                foreach (FilterRule rule in Rules.Where(r => r.Mode == FilterModes.Pattern)) rule.Overrides.Clear();
            Values.Clear();
            int index = 0;
            foreach (ScheduleLevelRow level in _levels)
            {
                index++;
                foreach (FilterRule rule in Rules)
                    Values.Add(new FilterValueRow { Rule = rule, Mode = rule.Mode, Level = level, Value = ScheduleService.ResolveText(rule, level, index) });
            }
            _builtRules = Fingerprint();
        }

        private void Commit()
        {
            foreach (DataGrid grid in new[] { RulesGrid, ValuesGrid })
            {
                grid.CommitEdit(DataGridEditingUnit.Cell, true);
                grid.CommitEdit(DataGridEditingUnit.Row, true);
            }
        }

        private void BeginValueEdit(object sender, DataGridBeginningEditEventArgs args)
        {
            var row = args.Row.Item as FilterValueRow;
            if (row == null || !row.CanEdit || _builtRules != Fingerprint())
                args.Cancel = true;
        }
        private void RebuildValues(object sender, RoutedEventArgs args)
        {
            Commit();
            try { BuildValues(true); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "KPLN: фильтры"); }
        }

        private void Apply(object sender, RoutedEventArgs args)
        {
            Commit();
            try
            {
                if (_builtRules != Fingerprint())
                    throw new InvalidOperationException("Правила изменились. Нажмите «Пересчитать по этажам», проверьте новые значения и затем примените настройки.");
                foreach (FilterValueRow value in Values.Where(v => v.CanEdit))
                {
                    int index = _levels.ToList().IndexOf(value.Level) + 1;
                    string automatic = ScheduleService.ResolveText(value.Rule, value.Level, index, false);
                    if (value.Rule.Mode == FilterModes.Manual || value.Value != automatic)
                        value.Rule.Overrides[value.Level.Key] = value.Value ?? "";
                    else value.Rule.Overrides.Remove(value.Level.Key);
                }
                int rowIndex = 0;
                foreach (ScheduleLevelRow level in _levels)
                {
                    rowIndex++;
                    foreach (FilterRule rule in Rules) ScheduleService.ResolveFilter(_document, _source.Schedule, rule, level, rowIndex);
                }
                foreach (FilterRule rule in Rules)
                {
                    FilterRule target = _source.Filters.Single(r => r.Index == rule.Index);
                    target.Mode = rule.Mode;
                    target.Pattern = rule.Pattern;
                    target.Overrides.Clear();
                    foreach (var pair in rule.Overrides) target.Overrides.Add(pair.Key, pair.Value);
                }
                DialogResult = true;
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "KPLN: фильтры"); }
        }
    }
}

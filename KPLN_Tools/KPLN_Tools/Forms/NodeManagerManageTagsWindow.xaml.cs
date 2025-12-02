using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class NodeManagerManageTagsWindow : Window
    {
        public ObservableCollection<string> AllTags { get; }
        public ObservableCollection<string> CurrentTags { get; }

        public List<string> ResultTags { get; private set; }

        public NodeManagerManageTagsWindow(
            IEnumerable<string> allTags,
            IEnumerable<string> currentTags)
        {
            InitializeComponent();

            AllTags = new ObservableCollection<string>(
                (allTags ?? Enumerable.Empty<string>())
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Select(t => t.Trim())
                    .Distinct(StringComparer.InvariantCultureIgnoreCase)
                    .OrderBy(t => t.Equals("DWG", StringComparison.InvariantCultureIgnoreCase) ? 0 : 1)
                    .ThenBy(t => t, StringComparer.CurrentCultureIgnoreCase));

            CurrentTags = new ObservableCollection<string>(
                (currentTags ?? Enumerable.Empty<string>())
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Select(t => t.Trim())
                    .Distinct(StringComparer.InvariantCultureIgnoreCase)
                    .OrderBy(t => t.Equals("DWG", StringComparison.InvariantCultureIgnoreCase) ? 0 : 1)
                    .ThenBy(t => t, StringComparer.CurrentCultureIgnoreCase));

            DataContext = this;
        }

        private void AddTag_Click(object sender, RoutedEventArgs e)
        {
            var selected = AllTagsCombo.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selected))
                return;

            if (!CurrentTags.Any(t => t.Equals(selected, StringComparison.InvariantCultureIgnoreCase)))
            {
                CurrentTags.Add(selected.Trim());
                ResortCurrentTags();
            }
        }

        private void AddCustomTag_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new NodeManagerInputTagWindow
            {
                Owner = this
            };

            if (dlg.ShowDialog() != true)
                return;

            var newTag = dlg.TagText;
            if (string.IsNullOrWhiteSpace(newTag))
                return;

            newTag = newTag.Trim();

            if (!CurrentTags.Any(t => t.Equals(newTag, StringComparison.InvariantCultureIgnoreCase)))
            {
                CurrentTags.Add(newTag);
                ResortCurrentTags();
            }

            if (!AllTags.Any(t => t.Equals(newTag, StringComparison.InvariantCultureIgnoreCase)))
            {
                AllTags.Add(newTag);
                ResortAllTags();
            }
        }

        private void DeleteTag_Click(object sender, RoutedEventArgs e)
        {
            var toRemove = CurrentTagsListBox.SelectedItems.Cast<string>().ToList();
            if (toRemove.Count == 0)
                return;

            foreach (var tag in toRemove)
            {
                var existing = CurrentTags.FirstOrDefault(t => t.Equals(tag, StringComparison.InvariantCultureIgnoreCase));
                if (existing != null)
                    CurrentTags.Remove(existing);
            }

            ResortCurrentTags();
        }

        private void ResortCurrentTags()
        {
            var sorted = CurrentTags
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .OrderBy(t => t.Equals("DWG", StringComparison.InvariantCultureIgnoreCase) ? 0 : 1)
                .ThenBy(t => t, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            CurrentTags.Clear();
            foreach (var t in sorted)
                CurrentTags.Add(t);
        }

        private void ResortAllTags()
        {
            var sorted = AllTags
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .OrderBy(t => t.Equals("DWG", StringComparison.InvariantCultureIgnoreCase) ? 0 : 1)
                .ThenBy(t => t, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            AllTags.Clear();
            foreach (var t in sorted)
                AllTags.Add(t);
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            ResultTags = CurrentTags.ToList();
            DialogResult = true;
            Close();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Forms;

namespace KPLN_Tools.Forms
{
    public partial class FilterByTagsWindowNodeManager : Window
    {
        public ObservableCollection<string> AvailableTags { get; }
        public ObservableCollection<string> SelectedTags { get; }

        public List<string> ResultTags => SelectedTags.ToList();

        public FilterByTagsWindowNodeManager(IEnumerable<string> allTags)
        {
            InitializeComponent();

            AvailableTags = new ObservableCollection<string>(
                allTags
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Distinct(StringComparer.InvariantCultureIgnoreCase)
                    .OrderBy(t => t)
            );

            SelectedTags = new ObservableCollection<string>();

            DataContext = this;
        }

        private void BtnAddTag_Click(object sender, RoutedEventArgs e)
        {
            var tag = TagsComboBox.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(tag))
                return;

            if (SelectedTags.Count >= 6)
            {
                System.Windows.MessageBox.Show("Можно выбрать не более 6 тегов.", "Фильтр по тегам", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (SelectedTags.Any(t =>
                    string.Equals(t, tag, StringComparison.InvariantCultureIgnoreCase)))
            {
                return;
            }

            SelectedTags.Add(tag);
        }

        private void RemoveTagButton_Click(object sender, RoutedEventArgs e)
        {
            var tag = (sender as FrameworkElement)?.DataContext as string;
            if (tag == null)
                return;

            SelectedTags.Remove(tag);
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedTags.Count == 0)
            {
                System.Windows.MessageBox.Show("Выберите хотя бы один тег.", "Фильтр по тегам", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

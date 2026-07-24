using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    internal partial class RevisionCloudSheetSelectionForm : Window
    {
        private const string ExcludedSheetsFolderName =
            "Не отображаются в текущей организации";

        private readonly ObservableCollection<RevisionCloudSheetNode> _rootNodes;
        private readonly int _sheetsCount;

        internal RevisionCloudSheetSelectionForm(
            IEnumerable<ViewSheet> sheets,
            BrowserOrganization browserOrganization)
        {
            InitializeComponent();

            List<ViewSheet> orderedSheets = sheets
                .OrderBy(sheet => sheet.SheetNumber, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(sheet => sheet.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            _sheetsCount = orderedSheets.Count;
            _rootNodes = BuildTree(orderedSheets, browserOrganization);

            foreach (RevisionCloudSheetNode node in EnumerateNodes(_rootNodes))
                node.SelectionChanged += Node_SelectionChanged;

            SheetsTree.ItemsSource = _rootNodes;
            UpdateSelectionSummary();
        }

        internal IReadOnlyList<ViewSheet> SelectedSheets { get; private set; }
            = new List<ViewSheet>();

        private static ObservableCollection<RevisionCloudSheetNode> BuildTree(
            IEnumerable<ViewSheet> sheets,
            BrowserOrganization browserOrganization)
        {
            ObservableCollection<RevisionCloudSheetNode> roots =
                new ObservableCollection<RevisionCloudSheetNode>();

            foreach (ViewSheet sheet in sheets)
            {
                IList<string> folderPath = GetFolderPath(sheet, browserOrganization);
                ObservableCollection<RevisionCloudSheetNode> currentLevel = roots;
                RevisionCloudSheetNode parent = null;

                foreach (string folderName in folderPath)
                {
                    RevisionCloudSheetNode folder = currentLevel
                        .FirstOrDefault(node =>
                            node.IsFolder
                            && string.Equals(
                                node.DisplayName,
                                folderName,
                                StringComparison.CurrentCultureIgnoreCase));

                    if (folder == null)
                    {
                        folder = RevisionCloudSheetNode.CreateFolder(folderName, parent);
                        currentLevel.Add(folder);
                    }

                    parent = folder;
                    currentLevel = folder.Children;
                }

                currentLevel.Add(RevisionCloudSheetNode.CreateSheet(sheet, parent));
            }

            SortNodes(roots);
            return roots;
        }

        private static IList<string> GetFolderPath(
            ViewSheet sheet,
            BrowserOrganization browserOrganization)
        {
            if (browserOrganization == null)
                return new List<string>();

            try
            {
                if (!browserOrganization.AreFiltersSatisfied(sheet.Id))
                    return new List<string> { ExcludedSheetsFolderName };

                return browserOrganization
                    .GetFolderItems(sheet.Id)
                    .Select(item =>
                        string.IsNullOrWhiteSpace(item.Name)
                            ? "(Без значения)"
                            : item.Name)
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        private static void SortNodes(ObservableCollection<RevisionCloudSheetNode> nodes)
        {
            foreach (RevisionCloudSheetNode folder in nodes.Where(node => node.IsFolder))
                SortNodes(folder.Children);

            List<RevisionCloudSheetNode> sorted = nodes
                .OrderBy(node => node.IsFolder ? 0 : 1)
                .ThenBy(node => node.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            nodes.Clear();

            foreach (RevisionCloudSheetNode node in sorted)
                nodes.Add(node);
        }

        private static IEnumerable<RevisionCloudSheetNode> EnumerateNodes(
            IEnumerable<RevisionCloudSheetNode> nodes)
        {
            foreach (RevisionCloudSheetNode node in nodes)
            {
                yield return node;

                foreach (RevisionCloudSheetNode child in EnumerateNodes(node.Children))
                    yield return child;
            }
        }

        private IEnumerable<RevisionCloudSheetNode> EnumerateSheetNodes()
        {
            return EnumerateNodes(_rootNodes).Where(node => !node.IsFolder);
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SetAllNodesChecked(true);
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            SetAllNodesChecked(false);
        }

        private void SetAllNodesChecked(bool isChecked)
        {
            foreach (RevisionCloudSheetNode node in _rootNodes)
                node.IsChecked = isChecked;

            UpdateSelectionSummary();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            List<ViewSheet> selectedSheets = EnumerateSheetNodes()
                .Where(node => node.IsChecked == true)
                .Select(node => node.Sheet)
                .ToList();

            if (selectedSheets.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Выберите хотя бы один лист для обработки.",
                    "KPLN. Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            SelectedSheets = selectedSheets;
            DialogResult = true;
        }

        private void Node_SelectionChanged(object sender, EventArgs e)
        {
            UpdateSelectionSummary();
        }

        private void UpdateSelectionSummary()
        {
            int selectedCount = EnumerateSheetNodes().Count(node => node.IsChecked == true);
            SelectionSummaryText.Text =
                "Выбрано: " + selectedCount + " из " + _sheetsCount;
        }
    }

    internal sealed class RevisionCloudSheetNode : INotifyPropertyChanged
    {
        private bool? _isChecked;

        private RevisionCloudSheetNode(
            string displayName,
            ViewSheet sheet,
            RevisionCloudSheetNode parent)
        {
            DisplayName = displayName;
            Sheet = sheet;
            Parent = parent;
            _isChecked = true;
        }

        internal event EventHandler SelectionChanged;

        public event PropertyChangedEventHandler PropertyChanged;

        public string DisplayName { get; }

        internal ViewSheet Sheet { get; }

        internal RevisionCloudSheetNode Parent { get; }

        public ObservableCollection<RevisionCloudSheetNode> Children { get; }
            = new ObservableCollection<RevisionCloudSheetNode>();

        internal bool IsFolder
        {
            get { return Sheet == null; }
        }

        public bool IsExpanded
        {
            get { return true; }
        }

        public bool? IsChecked
        {
            get { return _isChecked; }
            set { SetIsChecked(value, true, true); }
        }

        internal static RevisionCloudSheetNode CreateFolder(
            string name,
            RevisionCloudSheetNode parent)
        {
            return new RevisionCloudSheetNode(name, null, parent);
        }

        internal static RevisionCloudSheetNode CreateSheet(
            ViewSheet sheet,
            RevisionCloudSheetNode parent)
        {
            string displayName = sheet.SheetNumber + " — " + sheet.Name;
            return new RevisionCloudSheetNode(displayName, sheet, parent);
        }

        private void SetIsChecked(
            bool? value,
            bool updateChildren,
            bool updateParent)
        {
            if (_isChecked == value)
                return;

            _isChecked = value;
            OnPropertyChanged("IsChecked");

            if (updateChildren && value.HasValue)
            {
                foreach (RevisionCloudSheetNode child in Children)
                    child.SetIsChecked(value, true, false);
            }

            if (updateParent && Parent != null)
                Parent.RefreshCheckedState();

            EventHandler handler = SelectionChanged;

            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        private void RefreshCheckedState()
        {
            bool allChecked = Children.All(child => child.IsChecked == true);
            bool allUnchecked = Children.All(child => child.IsChecked == false);
            bool? value = allChecked ? true : allUnchecked ? (bool?)false : null;

            SetIsChecked(value, false, true);
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

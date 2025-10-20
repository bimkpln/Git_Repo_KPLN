using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Application = Autodesk.Navisworks.Api.Application;

namespace KPLN_Quantificator.Forms
{
    public enum SortKey
    {
        None,
        DisplayName,
        StatusThenAlpha,
        Level,
        GridIntersection,
        CreatedTime,
        ApprovedBy,
        AssignedTo,
        Distance,
    }

    public class BoolInverseToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool b = value is bool bb && bb;
            return b ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (value is Visibility v && v == Visibility.Collapsed);
        }
    }

    public partial class SortAllClashesForm : Window, INotifyPropertyChanged
    {
        private const int MaxRows = 8;

        public ObservableCollection<KeyValuePair<SortKey, string>> SortOptions { get; } =
            new ObservableCollection<KeyValuePair<SortKey, string>>(new[]
            {
                new KeyValuePair<SortKey, string>(SortKey.None,            "Выберите значение"),
                new KeyValuePair<SortKey, string>(SortKey.DisplayName,      "Имя"),
                new KeyValuePair<SortKey, string>(SortKey.StatusThenAlpha,  "Статус"),
                new KeyValuePair<SortKey, string>(SortKey.Level,            "Уровень"),
                new KeyValuePair<SortKey, string>(SortKey.GridIntersection, "Пересечение сетки"),
                new KeyValuePair<SortKey, string>(SortKey.CreatedTime,      "Найдено"),
                new KeyValuePair<SortKey, string>(SortKey.ApprovedBy,       "Кем утверждено"),
                new KeyValuePair<SortKey, string>(SortKey.AssignedTo,       "Утверждено"),
                new KeyValuePair<SortKey, string>(SortKey.Distance,         "Расстояние"),               
            });
     
        public ObservableCollection<SortRow> Rows { get; } = new ObservableCollection<SortRow>();

        public string AddHint
        {
            get
            {
                int left = MaxRows - Rows.Count;
                if (left <= 0) return $"Лимит достигнут: {MaxRows} строк.";
                return $"Можно добавить ещё {left} {WordForm(left, "строку", "строки", "строк")}.";
            }
        }

        public ICommand RemoveRowCmd { get; }

        public SortAllClashesForm()
        {
            InitializeComponent();
            DataContext = this;

            RemoveRowCmd = new RelayCommand<SortRow>(OnRemoveRow);

            Rows.Add(new SortRow { Key = SortKey.GridIntersection, IsProtected = true });
            Rows.Add(new SortRow { Key = SortKey.Level });

            Rows.CollectionChanged += Rows_CollectionChanged;

            UpdateButtonsState();
            OnPropertyChanged(nameof(AddHint));
        }

        private void Rows_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateButtonsState();
            OnPropertyChanged(nameof(AddHint));
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (Rows.Count >= MaxRows) return;

            var defaultKey = SortKey.DisplayName;
            if (Rows.Any(r => r.Key == defaultKey))
            {
                var firstFree = SortOptions.Select(o => o.Key)
                                           .FirstOrDefault(k => Rows.All(r => r.Key != k));
                if (!firstFree.Equals(default(SortKey)))
                    defaultKey = firstFree;
            }

            Rows.Add(new SortRow { Key = SortKey.None });
        }

        private void OnRemoveRow(SortRow row)
        {
            if (row == null || row.IsProtected) return;
            if (Rows.Count <= 1) return;
            Rows.Remove(row);
        }

        private void UpdateButtonsState()
        {
            if (BtnAdd != null) BtnAdd.IsEnabled = Rows.Count < MaxRows;
            if (BtnApplyAll != null) BtnApplyAll.IsEnabled = Rows.Count > 0;
        }

        private static string WordForm(int n, string one, string two, string many)
        {
            n = Math.Abs(n) % 100;
            int n1 = n % 10;
            if (n > 10 && n < 20) return many;
            if (n1 > 1 && n1 < 5) return two;
            if (n1 == 1) return one;
            return many;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public class SortRow
        {
            public SortKey Key { get; set; }
            public bool IsProtected { get; set; }
        }

        public class RelayCommand<T> : ICommand
        {
            private readonly Action<T> _execute;
            private readonly Predicate<T> _canExecute;

            public RelayCommand(Action<T> execute, Predicate<T> canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public bool CanExecute(object parameter) =>
                _canExecute?.Invoke((T)parameter) ?? true;

            public void Execute(object parameter) =>
                _execute((T)parameter);

            public event EventHandler CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnApplyAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var keysTopToBottom = Rows.Select(r => r.Key).ToList();
                ApplyToAllTests(keysTopToBottom);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Ошибка применения сортировки: {ex.Message}",
                    "Clash Detective", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ApplyToAllTests(IList<SortKey> keysTopToBottom)
        {
            var doc = Application.ActiveDocument ?? Application.MainDocument;
            if (doc == null || doc.IsClear) return;

            var dct = doc.GetClash()?.TestsData; 
            if (dct == null) return;

            var allTestGuids = dct.Tests.OfType<ClashTest>().Select(t => t.Guid).ToList();
            if (allTestGuids.Count == 0) return;

            using (var tr = doc.BeginTransaction("KPLN. Сортировка Clash Test"))
            {
                foreach (var testGuid in allTestGuids)
                {
                    var test = dct.Tests.OfType<ClashTest>().FirstOrDefault(t => t.Guid == testGuid);
                    if (test == null) continue;

   
                    for (int i = 0; i < keysTopToBottom.Count; i++)
                    {
                        ApplySingleKey(dct, test, keysTopToBottom[i]);
                        test = dct.Tests.OfType<ClashTest>().FirstOrDefault(t => t.Guid == testGuid);
                        if (test == null) break;
                    }
                }
                tr.Commit();
            }
        }

        private void ApplySingleKey(DocumentClashTests dct, ClashTest test, SortKey key)
        {
            switch (key)
            {
                case SortKey.None:
                    break;

                case SortKey.StatusThenAlpha:
                    SortStatusThenAlpha(dct, test);
                    break;

                case SortKey.DisplayName:
                    dct.TestsSortResults(test, ClashResultSortMode.DisplayNameSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.GridIntersection:
                    dct.TestsSortResults(test, ClashResultSortMode.GridIntersectionSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.Level:
                    dct.TestsSortResults(test, ClashResultSortMode.GridLevelSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.CreatedTime:
                    dct.TestsSortResults(test, ClashResultSortMode.CreatedTimeSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.ApprovedBy:
                    dct.TestsSortResults(test, ClashResultSortMode.ApprovedBySort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.AssignedTo:
                    dct.TestsSortResults(test, ClashResultSortMode.AssignedToSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                case SortKey.Distance:
                    dct.TestsSortResults(test, ClashResultSortMode.DistanceSort,
                                         ClashSortDirection.SortAscending, null);
                    break;

                default:
                    break;
            }
        }

        private void SortStatusThenAlpha(DocumentClashTests dct, ClashTest test)
        {
            var parent = test as GroupItem;
            if (parent == null) return;

            var all = parent.Children.OfType<SavedItem>().ToList();
            if (all.Count <= 1) return;

            var statusOrder = new[]
            {
                ClashResultStatus.New,
                ClashResultStatus.Active,
                ClashResultStatus.Reviewed,
                ClashResultStatus.Approved,
                ClashResultStatus.Resolved
            };

            var desired = new List<Guid>(all.Count);
            foreach (var st in statusOrder)
            {
                var bucket = all.Where(si => StatusOf(si) == st)
                                .OrderBy(si => NameOf(si), Comparer<string>.Create(NaturalCompare))
                                .ToList();

                foreach (var si in bucket)
                {
                    var g = si as ClashResultGroup;
                    if (g != null) { desired.Add(g.Guid); continue; }

                    var r = si as ClashResult;
                    if (r != null) desired.Add(r.Guid);
                }
            }

            for (int i = 0; i < desired.Count; i++)
            {
                var need = desired[i];
                test = dct.Tests.OfType<ClashTest>().FirstOrDefault(t => t.Guid == test.Guid);
                parent = test as GroupItem;
                if (parent == null) break;

                var children = parent.Children.OfType<SavedItem>().ToList();

                var curAtI = children.ElementAtOrDefault(i);
                bool already =
                    (curAtI as ClashResult)?.Guid == need ||
                    (curAtI as ClashResultGroup)?.Guid == need;
                if (already) continue;

                int j = children.FindIndex(si =>
                    (si as ClashResult)?.Guid == need ||
                    (si as ClashResultGroup)?.Guid == need);

                if (j < 0 || j == i) continue;

                dct.TestsMove(parent, j, parent, i);
            }
        }

        private static string NameOf(SavedItem si)
        {
            var g = si as ClashResultGroup;
            if (g != null) return g.DisplayName ?? "";
            var r = si as ClashResult;
            if (r != null) return r.DisplayName ?? "";
            return si != null ? (si.DisplayName ?? "") : "";
        }

        private static ClashResultStatus StatusOf(SavedItem si)
        {
            var r = si as ClashResult;
            if (r != null) return r.Status;
            var g = si as ClashResultGroup;
            if (g != null) return g.Status;
            return ClashResultStatus.New;
        }

        private static int NaturalCompare(string a, string b)
        {
            if (a == null) a = "";
            if (b == null) b = "";

            var rx = new Regex(@"\d+");
            var asplit = rx.Split(a);
            var bsplit = rx.Split(b);

            var anum = rx.Matches(a).Cast<Match>().Select(m => long.Parse(m.Value)).ToArray();
            var bnum = rx.Matches(b).Cast<Match>().Select(m => long.Parse(m.Value)).ToArray();

            int i = 0, j = 0, ni = 0, nj = 0;
            for (; i < asplit.Length && j < bsplit.Length; i++, j++)
            {
                int cmp = string.Compare(asplit[i], bsplit[j], StringComparison.CurrentCultureIgnoreCase);
                if (cmp != 0) return cmp;

                if (ni < anum.Length && nj < bnum.Length)
                {
                    if (anum[ni] != bnum[nj]) return anum[ni].CompareTo(bnum[nj]);
                    ni++; nj++;
                }
            }
            return asplit.Length - bsplit.Length;
        }
    }
}
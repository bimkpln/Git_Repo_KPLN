using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;

namespace KPLN_Quantificator.Forms
{
    public partial class StatusChangeClash : Window
    {
        private const string None = "Не выбрано";

        private static readonly (string Ru, ClashResultStatus Api)[] StatusMap =
        {
            ("Создать", (ClashResultStatus)0),
            ("Активные", (ClashResultStatus)1),
            ("Проверенные", (ClashResultStatus)2),
            ("Подтвержденные", (ClashResultStatus)3),
            ("Исправленные", (ClashResultStatus)4),
        };

        public StatusChangeClash()
        {
            InitializeComponent();

            var items = new List<string> { None };
            items.AddRange(StatusMap.Select(x => x.Ru));

            cbOldStatus.ItemsSource = items;
            cbNewStatus.ItemsSource = items;

            cbOldStatus.SelectedItem = None;
            cbNewStatus.SelectedItem = None;
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            var fromRu = cbOldStatus.SelectedItem as string;
            var toRu = cbNewStatus.SelectedItem as string;

            var fromNone = string.IsNullOrWhiteSpace(fromRu) || fromRu == None;
            var toNone = string.IsNullOrWhiteSpace(toRu) || toRu == None;

            if (fromNone && toNone)
            {
                MessageBox.Show("Не выбраны оба статуса.", "KPLN", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (fromNone)
            {
                MessageBox.Show("Не выбран \"Статус был\".", "KPLN", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (toNone)
            {
                MessageBox.Show("Не выбран \"Статус станет\".", "KPLN", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (fromRu == toRu)
            {
                MessageBox.Show("Статусы совпадают: менять нечего.", "KPLN", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var fromApi = StatusMap.First(x => x.Ru == fromRu).Api;
            var toApi = StatusMap.First(x => x.Ru == toRu).Api;

            int changed = ChangeStatusesInAllTests_WithApproveFields(fromApi, toApi, Environment.UserName, DateTime.Now, clearApproveFieldsWhenLeavingApproved: true);

            MessageBox.Show($"Готово. Изменено коллизий: {changed}", "KPLN", MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private static int ChangeStatusesInAllTests_WithApproveFields(
            ClashResultStatus from, ClashResultStatus to, string approvedBy, DateTime approvedTime, bool clearApproveFieldsWhenLeavingApproved)
        {
            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument ?? Autodesk.Navisworks.Api.Application.MainDocument;
            if (doc == null || doc.IsClear) return 0;

            var clash = DocumentExtensions.GetClash(doc);
            var dct = clash != null ? clash.TestsData : null;
            if (dct?.Tests == null || dct.Tests.Count == 0) return 0;

            int changed = 0;

            using (var tr = doc.BeginTransaction("KPLN: Смена статуса"))
            {
                foreach (var test in EnumerateClashTests(dct.Tests))
                {
                    foreach (var cr in EnumerateClashResultsOnly(test))
                    {
                        if (cr == null) continue;
                        if (cr.Status != from) continue;

                        dct.TestsEditResultStatus((IClashResult)cr, to);
                        changed++;


                        if ((int)to == 3)
                        {
                            Try(() => dct.TestsEditResultApprovedBy((IClashResult)cr, approvedBy));
                            Try(() => dct.TestsEditResultApprovedTime((IClashResult)cr, approvedTime));
                        }

                        if (clearApproveFieldsWhenLeavingApproved && (int)from == 3 && (int)to != 3)
                        {
                            Try(() => dct.TestsEditResultApprovedBy((IClashResult)cr, ""));
                            Try(() => dct.TestsEditResultApprovedTime((IClashResult)cr, DateTime.MinValue));
                        }
                    }
                }

                tr.Commit();
            }

            RefreshClashUI_NoRecalc(dct);
            return changed;
        }

        private static IEnumerable<ClashTest> EnumerateClashTests(SavedItemCollection root)
        {
            if (root == null) yield break;

            foreach (SavedItem i in root)
            {
                var t = i as ClashTest;
                if (t != null) yield return t;

                var g = i as GroupItem;
                if (g?.Children == null) continue;

                foreach (var childTest in EnumerateClashTests(g.Children))
                    yield return childTest;
            }
        }

        private static IEnumerable<ClashResult> EnumerateClashResultsOnly(ClashTest test)
        {
            if (test == null) yield break;

            var g = test as GroupItem;
            if (g?.Children == null) yield break;

            foreach (SavedItem i in g.Children)
            {
                foreach (var cr in EnumerateClashResultsOnly_Tree(i))
                    yield return cr;
            }
        }

        private static IEnumerable<ClashResult> EnumerateClashResultsOnly_Tree(SavedItem item)
        {
            if (item == null) yield break;

            var cr = item as ClashResult;
            if (cr != null) yield return cr;

            var g = item as GroupItem;
            if (g?.Children == null) yield break;

            foreach (SavedItem c in g.Children)
            {
                foreach (var nested in EnumerateClashResultsOnly_Tree(c))
                    yield return nested;
            }
        }

        private static void RefreshClashUI_NoRecalc(DocumentClashTests dct)
        {
            if (dct?.Tests == null || dct.Tests.Count == 0) return;

            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument ?? Autodesk.Navisworks.Api.Application.MainDocument;
            if (doc == null || doc.IsClear) return;

            Try(() => System.Windows.Forms.Application.DoEvents());

            using (var tr = doc.BeginTransaction("KPLN: Обновление UI"))
            {
                foreach (var test in EnumerateClashTests(dct.Tests))
                {
                    Try(() =>
                    {
                        var copy = (test as SavedItem)?.CreateCopy() as ClashTest;
                        if (copy != null)
                            dct.TestsEditTestFromCopy(test, copy);
                    });
                }

                tr.Commit();
            }

            Try(() => System.Windows.Forms.Application.DoEvents());
        }

        private static void Try(Action a)
        {
            try { a(); } catch { }
        }
    }
}
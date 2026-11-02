using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.Services.SQLite;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace KPLN_Clashes_Ribbon.Forms
{
    public partial class ReportForm
    {
        private readonly string _reportSessionId = Guid.NewGuid().ToString("N");
        private readonly Dictionary<int, long> _rowRevisions = new Dictionary<int, long>();
        private Dictionary<int, ReportItem> _itemsById;
        private Dictionary<int, string> _departmentNames;
        private readonly DispatcherTimer _changeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(20) };
        private long _lastReportRevision = -1;
        private int _localSaveGeneration;
        private bool _isPolling;
        private bool _isEditing;
        private bool _isClosed;
        private bool _liveCheckFailed;
        private readonly Dictionary<int, string> _pendingNameRows = new Dictionary<int, string>();
        private Task _nameRefreshTask = Task.CompletedTask;
        private Task _metadataRefreshTask = Task.CompletedTask;
        private MetadataUpdate _pendingMetadata;
        private bool _metadataNeedsSummary;

        private sealed class MetadataUpdate
        {
            public int ReportId;
            public int GroupId;
            public long Revision;
            public string UserName;
            public string FullName;
        }

        private void StartLiveUpdates()
        {
            InitializeLiveCaches();
            Loaded += LiveUpdatesLoaded;
            Closed += LiveUpdatesClosed;
            _changeTimer.Tick += LiveUpdatesTick;
        }

        private void InitializeLiveCaches()
        {
            _itemsById = ReportInstancesColl.SelectMany(item => item.SubElements.Concat(new[] { item }))
                .GroupBy(item => item.Id).ToDictionary(group => group.Key, group => group.First());
            _departmentNames = _itemsById.Values.SelectMany(item => item.SubDepartmentBtns)
                .Where(button => !string.IsNullOrWhiteSpace(button.Name))
                .GroupBy(button => button.Id).ToDictionary(group => group.Key, group => group.First().Name);
        }

        private async void LiveUpdatesLoaded(object sender, RoutedEventArgs e)
        {
            Loaded -= LiveUpdatesLoaded;
            await PollChangesAsync();
        }

        private async void LiveUpdatesTick(object sender, EventArgs e) => await PollChangesAsync();

        private void LiveUpdatesClosed(object sender, EventArgs e)
        {
            _isClosed = true;
            _pendingNameRows.Clear();
            _changeTimer.Stop();
            _changeTimer.Tick -= LiveUpdatesTick;
            Loaded -= LiveUpdatesLoaded;
            Closed -= LiveUpdatesClosed;
        }

        private bool ShowCommentInput(UserTextInput input)
        {
            _isEditing = true;
            try
            {
                return input.ShowDialog() == true;
            }
            finally
            {
                _isEditing = false;
            }
        }

        private async Task PollChangesAsync(bool force = false)
        {
            if (_isClosed || _isPolling)
                return;

            _changeTimer.Stop();
            if (_isSaving || _isEditing)
            {
                _changeTimer.Start();
                return;
            }

            _isPolling = true;
            int saveGeneration = _localSaveGeneration;
            try
            {
                var snapshot = await _sqliteService_ReportInstanceDB.ReadChangesAsync(_lastReportRevision, force);

                // Ввод мог начаться во время запроса. Курсор не двигаем: изменения прочитаем снова.
                if (_isClosed || _isSaving || _isEditing || saveGeneration != _localSaveGeneration)
                    return;

                ApplyLiveRows(snapshot.Items, force);
                QueueCommentNames(snapshot.Items);
                _lastReportRevision = snapshot.Revision;
                string notice = BuildChangeNotice(snapshot.Changes, _reportSessionId, _itemsById, _departmentNames);
                if (!string.IsNullOrEmpty(notice))
                {
                    LiveChangesText.Text = DateTime.Now.ToString("HH:mm:ss") + " — " + notice;
                    LiveChangesBanner.Visibility = Visibility.Visible;
                }
                if (_liveCheckFailed)
                    LiveCheckStatusText.Text = "Связь восстановлена. Проверка каждые 20 секунд.";
                _liveCheckFailed = false;
            }
            catch (Exception)
            {
                if (!_isClosed)
                {
                    _liveCheckFailed = true;
                    LiveCheckStatusText.Text = "Не удалось проверить изменения. Повтор через 20 секунд.";
                }
            }
            finally
            {
                _isPolling = false;
                if (!_isClosed)
                    _changeTimer.Start();
            }
        }

        private void QueueCommentNames(SQLiteService_ReportItemsDB.SavedItem[] rows)
        {
            if (_isClosed)
                return;
            foreach (var row in rows)
                _pendingNameRows[row.Id] = row.Comments;
            if (_pendingNameRows.Count > 0 && _nameRefreshTask.IsCompleted)
                _nameRefreshTask = RefreshCommentNamesAsync();
        }

        private async Task RefreshCommentNamesAsync()
        {
            while (_pendingNameRows.Count > 0 && !_isClosed)
            {
                var batch = _pendingNameRows.ToArray();
                _pendingNameRows.Clear();
                try
                {
                    await ReportItemComment.LoadNamesForCommentsAsync(batch.Select(row => row.Value));
                    if (_isClosed)
                        return;

                    // Только подписи авторов: данные строк и пользовательский ввод не заменяем.
                    foreach (var row in batch)
                    {
                        if (_itemsById.TryGetValue(row.Key, out ReportItem item))
                            foreach (ReportItemComment comment in item.CommentCollection)
                                comment.RefreshUserName();
                    }
                }
                catch (Exception ex)
                {
                    // До следующей загрузки остаются сохраненные в комментариях учетные имена.
                    System.Diagnostics.Trace.TraceError("Clashes: загрузка имен авторов: " + ex);
                }
            }
        }

        private void QueueMetadataRefresh(SQLiteService_ReportItemsDB.SavedItem[] saved, bool refreshStatus,
            int reportId, int groupId, string userName, string fullName)
        {
            _pendingMetadata = new MetadataUpdate
            {
                ReportId = reportId, GroupId = groupId, Revision = saved.Max(row => row.Revision),
                UserName = userName, FullName = fullName
            };
            _metadataNeedsSummary |= refreshStatus;
            if (!_isClosed)
            {
                SaveFollowUpText.Text = "Изменения сохранены. Сводка обновляется…";
                SaveFollowUpText.ToolTip = null;
            }
            if (_metadataRefreshTask.IsCompleted)
                _metadataRefreshTask = RefreshMetadataAsync();
        }

        private async Task RefreshMetadataAsync()
        {
            // Один обработчик на окно. Пока он занят, следующие сохранения объединяются.
            while (_pendingMetadata != null)
            {
                MetadataUpdate update = _pendingMetadata;
                _pendingMetadata = null;
                bool refreshStatus = _metadataNeedsSummary;
                _metadataNeedsSummary = false;
                try
                {
                    var summary = refreshStatus ? await _sqliteService_ReportInstanceDB.ReadSummaryAsync() : null;
                    await _sqliteService_MainDB.UpdateBatchMetadataAsync(update.ReportId, update.GroupId,
                        update.Revision, summary, update.UserName, update.FullName);
                    if (!_isClosed && _pendingMetadata == null)
                        SaveFollowUpText.Text = string.Empty;
                }
                catch (Exception ex)
                {
                    // Комментарии уже сохранены. Следующая правка повторит только обновление сводки.
                    _metadataNeedsSummary |= refreshStatus;
                    if (!_isClosed && _pendingMetadata == null)
                    {
                        SaveFollowUpText.Text = "Изменения сохранены, но сводка не обновилась. Повторять правку не нужно.";
                        SaveFollowUpText.ToolTip = ex.Message;
                    }
                    System.Diagnostics.Trace.TraceError("Clashes: обновление сводки: " + ex);
                }
            }
        }

        private void DismissLiveChanges(object sender, RoutedEventArgs e)
        {
            LiveChangesBanner.Visibility = Visibility.Collapsed;
        }

        internal static string BuildChangeNotice(SQLiteService_ReportItemsDB.ReportChange[] changes, string sessionId,
            Dictionary<int, ReportItem> items, Dictionary<int, string> departments)
        {
            var externalChanges = changes.Where(c => c.SessionId != sessionId).ToArray();
            if (externalChanges.Length == 0)
                return string.Empty;

            return string.Join(Environment.NewLine, externalChanges
                .GroupBy(c => string.IsNullOrWhiteSpace(c.Actor) ? "Другой пользователь" : c.Actor)
                .Select(group =>
                {
                    // Берем действие из события, а не из конечного статуса строки:
                    // за один интервал разные люди могли закрыть и снова открыть одно замечание.
                    var actions = group.Where(c => c.StatusChanged != 0)
                        .Select(c => new { c.ItemId, Action = GetChangeAction(c, departments) })
                        .Concat(group.Where(c => c.CommentsChanged != 0 && c.StatusChanged == 0)
                            .Select(c => new { c.ItemId, Action = "обновил комментарии" }));
                    var parts = actions.GroupBy(action => action.Action).Select(action =>
                    {
                        int[] ids = action.Select(entry => entry.ItemId).Distinct().ToArray();
                        string names = string.Join(", ", ids.Take(3).Select(id => GetNoticeItemName(id, items)));
                        if (ids.Length > 3)
                            names += ", ещё " + (ids.Length - 3);
                        return action.Key + " (" + ids.Length + "): " + names;
                    });
                    return group.Key + " — " + string.Join("; ", parts);
                }));
        }

        private static string GetChangeAction(SQLiteService_ReportItemsDB.ReportChange change, Dictionary<int, string> departments)
        {
            // До обновления журнала конечный статус события не сохранялся.
            if (!change.StatusId.HasValue)
                return "изменил статус";

            switch ((ClashesMainCollection.KPItemStatus)change.StatusId.Value)
            {
                case ClashesMainCollection.KPItemStatus.Closed: return "закрыл";
                case ClashesMainCollection.KPItemStatus.Approved: return "допустил";
                case ClashesMainCollection.KPItemStatus.Opened: return "вернул в работу";
                case ClashesMainCollection.KPItemStatus.New: return "перевёл в новые";
                case ClashesMainCollection.KPItemStatus.Delegated:
                    int departmentId = change.DelegatedDepartmentId ?? -1;
                    if (departments.TryGetValue(departmentId, out string name))
                        return "делегировал отделу " + name;
                    return departmentId > 0 ? "делегировал отделу №" + departmentId : "делегировал";
                default: return "изменил статус";
            }
        }

        private static string GetNoticeItemName(int id, Dictionary<int, ReportItem> items)
        {
            if (!items.TryGetValue(id, out ReportItem item) || string.IsNullOrWhiteSpace(item.Name))
                return "замечание №" + id;

            string name = item.Name.Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (name.Length > 80)
                name = name.Substring(0, 77) + "…";
            return "«" + name + "»";
        }

        private void ApplyLiveRows(SQLiteService_ReportItemsDB.SavedItem[] rows, bool refreshView = false)
        {
            if (rows.Length == 0 && !refreshView)
                return;

            var updates = rows.Where(row => _itemsById.ContainsKey(row.Id)
                && (!_rowRevisions.TryGetValue(row.Id, out long applied) || applied <= row.Revision)).ToArray();
            // Из изменяемых полей только статус участвует в фильтре отчета.
            bool refreshFilter = refreshView || (StatusFilterCBX.SelectedIndex > 0
                && updates.Any(row => _itemsById[row.Id].StatusId != row.StatusId));
            var selected = refreshFilter
                ? ReportControll.SelectedItems.Cast<ReportItem>().Select(item => item.Id).ToHashSet()
                : new HashSet<int>();
            ScrollViewer scroll = refreshFilter ? FindVisualChild<ScrollViewer>(ReportControll) : null;
            double vertical = scroll?.VerticalOffset ?? 0;
            double horizontal = scroll?.HorizontalOffset ?? 0;
            ListBoxItem anchor = scroll == null ? null : FindVisualChildren<ListBoxItem>(ReportControll)
                .FirstOrDefault(container => container.TransformToAncestor(scroll).Transform(new Point()).Y >= 0);
            object anchorItem = anchor?.DataContext;
            double anchorY = anchor == null ? 0 : anchor.TransformToAncestor(scroll).Transform(new Point()).Y;
            foreach (var row in updates)
            {
                ReportItem item = _itemsById[row.Id];
                _rowRevisions[row.Id] = row.Revision;
                string comments = row.Comments ?? string.Empty;
                string currentComments = string.Join(ClashesMainCollection.StringSeparatorItem, item.CommentCollection.Select(c => c.ToString()));
                if (item.StatusId != row.StatusId)
                    item.StatusId = row.StatusId;
                if (item.DelegatedDepartmentId != row.DelegatedDepartmentId)
                {
                    ResetDelegateBtnBrush(item);
                    item.DelegatedDepartmentId = row.DelegatedDepartmentId;
                    SubDepartmentBtn button = item.SubDepartmentBtns.FirstOrDefault(btn => btn.Id == row.DelegatedDepartmentId);
                    if (button != null)
                        SetDelegateBtnBrush(item, button);
                }
                if (currentComments != comments)
                    item.CommentCollection = ReportItemComment.ParseComments(comments, item);
            }

            if (!refreshFilter)
                return;

            // Коллекцию и сами ReportItem не заменяем: изображения, раскрытые строки и фильтры остаются.
            FilteredInstancesColl.Refresh();
            foreach (ReportItem item in ReportControll.Items)
            {
                if (selected.Contains(item.Id) && !ReportControll.SelectedItems.Contains(item))
                    ReportControll.SelectedItems.Add(item);
            }

            if (scroll != null)
            {
                ReportControll.UpdateLayout();
                var currentAnchor = anchorItem == null ? null : ReportControll.ItemContainerGenerator.ContainerFromItem(anchorItem) as ListBoxItem;
                if (currentAnchor != null)
                    vertical = scroll.VerticalOffset + currentAnchor.TransformToAncestor(scroll).Transform(new Point()).Y - anchorY;
                scroll.ScrollToVerticalOffset(vertical);
                scroll.ScrollToHorizontalOffset(horizontal);
            }
        }

        private static T FindVisualChild<T>(DependencyObject root) where T : DependencyObject => FindVisualChildren<T>(root).FirstOrDefault();

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject root) where T : DependencyObject
        {
            for (int index = 0; index < VisualTreeHelper.GetChildrenCount(root); index++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(root, index);
                if (child is T match)
                    yield return match;
                foreach (T descendant in FindVisualChildren<T>(child))
                    yield return descendant;
            }
        }
    }
}

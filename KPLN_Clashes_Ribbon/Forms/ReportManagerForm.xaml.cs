using HtmlAgilityPack;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.Services;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;
using static KPLN_Clashes_Ribbon.Tools.HTMLTools;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using Image = System.Windows.Controls.Image;
using Path = System.IO.Path;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportManagerForm.xaml
    /// </summary>
    public partial class ReportManagerForm : Window
    {
        private readonly DBProject _project;
        private readonly Services.SQLite.SQLiteService_MainDB _sqliteService_MainDB = new Services.SQLite.SQLiteService_MainDB();
        private string _reportNameTBxData = string.Empty;
        private string _reportBitrixIdTBxData = string.Empty;

        public ReportManagerForm(DBProject project)
        {
            _project = project ?? throw new ArgumentNullException("\n[KPLN]: Попытка передачи пустого проекта\n");

            InitializeComponent();
            DataContext = this;

            UpdateReportGroups();

            if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
                btnAddGroup.Visibility = Visibility.Visible;
            else
                btnAddGroup.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Отфильтрованная коллекция элементов, которая является контекстом для окна
        /// </summary>
        public ICollectionView FilteredRepGroupColl { get; private set; }

        /// <summary>
        /// Обновить выбранную группу
        /// </summary>
        public void UpdateSelectedReportGroup(ReportGroup repGroup)
        {
            if (!(FilteredRepGroupColl.SourceCollection is ReportGroup[] groups)) return;

            for (int i = 0; i < groups.Length; i++)
            {
                if (groups[i].Id == repGroup.Id)
                {
                    string searchText = groups[i].SearchText;
                    ReportGroup tempRG = _sqliteService_MainDB.GetReportGroup_ById(groups[i].Id);

                    groups[i] = tempRG;
                    groups[i].IsExpandedItem = true;
                    tempRG.SearchText = searchText;

                    FilteredRepGroupColl = CollectionViewSource.GetDefaultView(groups.Select(gr => SetReportsToReportGroup(gr)).ToArray());
                    FilteredRepGroupColl.Filter += FilterRepGroups;

                    iControllGroups.ItemsSource = FilteredRepGroupColl;

                    ApplySearchToReportGroup(groups[i]);

                    break;
                }
            }
        }

        /// <summary>
        /// Обновить коллекцию групп
        /// </summary>
        public void UpdateReportGroups()
        {
            ReportGroup[] groups = _sqliteService_MainDB
                .GetReportGroups_ByDBProject(_project)
                .OrderBy(gr => gr.Status != KPItemStatus.Closed)
                .ThenBy(gr => gr.Id)
                .ToArray();

            if (groups != null)
            {
                FilteredRepGroupColl = CollectionViewSource.GetDefaultView(groups.Select(gr => SetReportsToReportGroup(gr)).ToArray());
                FilteredRepGroupColl.Filter += FilterRepGroups;

                iControllGroups.ItemsSource = FilteredRepGroupColl;
            }
        }

        private ReportGroup SetReportsToReportGroup(ReportGroup group)
        {
            ObservableCollection<Report> reports = _sqliteService_MainDB.GetReports_ByReportGroupId(group.Id);

            // Настройка визуализации ReportGroup если по отчетам Report была активность
            if (group.Status == Core.ClashesMainCollection.KPItemStatus.New)
            {
                IEnumerable<Report> notNewReports = reports.Where(r => r.Status != Core.ClashesMainCollection.KPItemStatus.New);
                if (notNewReports.Any())
                {
                    group.Status = Core.ClashesMainCollection.KPItemStatus.Opened;
                    _sqliteService_MainDB.UpdateItemStatus_ByTableAndItemId(Core.ClashesMainCollection.KPItemStatus.Opened, MainDB_Enumerator.ReportGroups, group.Id);
                }
            }

            foreach (Report report in reports)
            {
                // Настройка визуализации Report если отчеты закрыты (смена картинки)
                if (group.Status != Core.ClashesMainCollection.KPItemStatus.Closed)
                    report.IsGroupEnabled = Visibility.Visible;
                else
                    report.IsGroupEnabled = Visibility.Collapsed;

                if (group.Reports.All(rep => rep.Id != report.Id))
                    group.Reports.Add(report);
            }

            return group;
        }

        private bool FilterRepGroups(object obj)
        {
            if (obj is ReportGroup rGroup)
            {
                if ((bool)this.ShowClosedReportGroups.IsChecked)
                    return CheckParamData(rGroup);
                else if (!(bool)this.ShowClosedReportGroups.IsChecked && rGroup.Status != KPItemStatus.Closed)
                    return CheckParamData(rGroup);
            }

            return false;
        }

        private bool CheckParamData(ReportGroup rGroup)
        {
            if (string.IsNullOrEmpty(_reportNameTBxData)
                    && string.IsNullOrEmpty(_reportBitrixIdTBxData))
                return true;

            string rGrName = rGroup.Name;


            if (!string.IsNullOrEmpty(_reportNameTBxData)
                && rGrName.ToLower().Contains(_reportNameTBxData))
                return true;

            string rGrBitrId_AR = rGroup.BitrixTaskIdAR.ToString();
            string rGrBitrId_KR = rGroup.BitrixTaskIdKR.ToString();
            string rGrBitrId_OV = rGroup.BitrixTaskIdOV.ToString();
            string rGrBitrId_ITP = rGroup.BitrixTaskIdITP.ToString();
            string rGrBitrId_VK = rGroup.BitrixTaskIdVK.ToString();
            string rGrBitrId_AUPT = rGroup.BitrixTaskIdAUPT.ToString();
            string rGrBitrId_EOM = rGroup.BitrixTaskIdEOM.ToString();
            string rGrBitrId_SS = rGroup.BitrixTaskIdSS.ToString();
            string rGrBitrId_AV = rGroup.BitrixTaskIdAV.ToString();

            if (!string.IsNullOrEmpty(_reportBitrixIdTBxData)
                && (rGrBitrId_AR.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_KR.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_OV.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_ITP.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_VK.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_AUPT.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_EOM.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_EOM.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_SS.Contains(_reportBitrixIdTBxData)
                    || rGrBitrId_AV.Contains(_reportBitrixIdTBxData)))
                return true;

            return false;
        }

        private void OnBtnRemoveReport(object sender, RoutedEventArgs args)
        {
            KPTaskDialog dialog = new KPTaskDialog(
                this,
                "Удалить отчет",
                "Необходимо подтверждение",
                "Вы уверены, что хотите удалить данный отчет?",
                Core.ClashesMainCollection.KPTaskDialogIcon.Question,
                true,
                "После удаления данные о статусе и комментарии будут безвозвратно потеряны!");
            dialog.ShowDialog();

            if (dialog.DialogResult == Core.ClashesMainCollection.KPTaskDialogResult.Ok)
            {
                if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
                {
                    Report report = (sender as System.Windows.Controls.Button).DataContext as Report;
                    _sqliteService_MainDB.DeleteReportAndReportItems_ByReportId(report);
                    UpdateSelectedReportGroup(report.ReportGroup);
                }
                else
                {
                    KPTaskDialog dialogError = new KPTaskDialog(
                        this,
                        "Удалить отчет",
                        "Ошибка удаления",
                        "Удалить может только сотрудник BIM-отдела",
                        Core.ClashesMainCollection.KPTaskDialogIcon.Question,
                        true);
                    dialogError.ShowDialog();
                }
            }
        }

        private void OnBtnAddReport(object sender, RoutedEventArgs args)
        {
            if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
            {
                ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                int repInstIndex = 0;
                try
                {
                    OpenFileDialog dialog = new OpenFileDialog
                    {
                        Filter = "html report (*.html)|*.html",
                        Title = "Выберите отчет(ы) NavisWorks в формате .html",
                        CheckFileExists = true,
                        CheckPathExists = true,
                        Multiselect = true
                    };
                    DialogResult result = dialog.ShowDialog();

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        int taskCount = 0;
                        Task[] riWorkerTasks = new Task[dialog.FileNames.Length];
                        foreach (string file_name in dialog.FileNames)
                        {
                            FileInfo file = new FileInfo(file_name);
                            string reportInstanceName = file.Name.Replace(".html", "");

                            // Публикую запись о новом репорте
                            FileInfo db_FI = GenerateNewPath_DBForReportInstance(group, ++repInstIndex);
                            int newReportId = _sqliteService_MainDB.PostReport_NewReport_ByNameAndReportGroup(reportInstanceName, group, db_FI);

                            ObservableCollection<ReportItem> reportInstances = ParseHtmlToRepInstColelction(file, group, newReportId);
                            if (reportInstances != null)
                            {
                                if (reportInstances.Count != 0)
                                {
                                    //Создаю БД под item и публикую в него данные
                                    Services.SQLite.SQLiteService_ReportItemsDB sqliteService_ReportItemsDB = new Services.SQLite.SQLiteService_ReportItemsDB(db_FI.FullName);
                                    Task fiWorkTask = Task.Run(() =>
                                    {
                                        sqliteService_ReportItemsDB.CreateDbFile_ByReports();
                                        sqliteService_ReportItemsDB.PostNewItems_ByReports(reportInstances);
                                    });
                                    riWorkerTasks[taskCount++] = fiWorkTask;
                                }
                                else
                                {
                                    string[] parts = new string[] { "«Изображение»",
                                        "«Наименование конфлика»",
                                        "«Расположение сетки»",
                                        "«Точка конфликта»",
                                        "«Идентификатор элемента (Элемент 1)»",
                                        "«Путь (Элемент 1)»",
                                        "«Идентификатор элемента (Элемент 2)»",
                                        "«Путь (Элемент 2)»" };
                                    Print(string.Format("Не удалось создать отчет на основе файла: «{0}»;\nУбедитесь, что файл является отчетом NavisWorks и содержит следующую информацию: {1}", file.Name, string.Join(", ", parts)),
                                        MessageType.Error);
                                }
                            }
                        }
                        Task.WaitAll(riWorkerTasks);

                        UpdateSelectedReportGroup(group);
                    }
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }
        }

        /// <summary>
        /// Генерация пути БД для хранения данных по ReportInstance
        /// </summary>
        private FileInfo GenerateNewPath_DBForReportInstance(ReportGroup group, int index)
        {
            // Проверка на наличие файла с таким именем.
            while (File.Exists(Path.Combine(@"Z:\Отдел BIM\03_Скрипты\08_Базы данных\Clashes_ReportItems", $"Clashes_RG_{group.Id}_RI_{index}.db")))
                index++;

            return new FileInfo(Path.Combine(@"Z:\Отдел BIM\03_Скрипты\08_Базы данных\Clashes_ReportItems", $"Clashes_RG_{group.Id}_RI_{index}.db"));
        }

        /// <summary>
        /// Разложить html-файл на коллекцию ReportInstance
        /// </summary>
        /// <param name="file">Html-файл для парсинга</param>
        /// <returns></returns>
        private ObservableCollection<ReportItem> ParseHtmlToRepInstColelction(FileInfo file, ReportGroup repGroup, int reportId)
        {
            ObservableCollection<ReportItem> reportInstances = new ObservableCollection<ReportItem>();
            if (file.Extension == ".html")
            {
                HtmlAgilityPack.HtmlDocument htmlSnippet = new HtmlAgilityPack.HtmlDocument();
                using (FileStream stream = file.OpenRead())
                {
                    htmlSnippet.Load(stream);
                }
                int num_id = 0;
                List<string> headers = new List<string>();
                int parentGroupId = -1;
                bool decode = false;
                foreach (HtmlNode node in htmlSnippet.DocumentNode.SelectNodes("//tr"))
                {
                    bool resetgroupid = false;
                    bool isGroupInstance = false;
                    bool addInstance = false;
                    string class_name = node.GetAttributeValue("class", "NONE");
                    //
                    string name = string.Empty;
                    string image = string.Empty;
                    string point = string.Empty;
                    string id1 = string.Empty;
                    string id2 = string.Empty;
                    string name1 = string.Empty;
                    string name2 = string.Empty;
                    string reportParentGroupComments = string.Empty;
                    string reportItemComments = string.Empty;
                    if (class_name == "headerRow" && headers.Count == 0)
                    {
                        if (IsMainHeader(node, out List<string> out_headers, out bool todecode))
                        {
                            headers = out_headers;
                            decode = todecode;
                        }
                    }
                    if (headers.Count == 0)
                        continue;

                    if (class_name == "clashGroupRow")
                    {
                        isGroupInstance = true;
                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id"), decode));
                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id", true), decode));
                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                        reportParentGroupComments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                        num_id++;
                        parentGroupId = num_id;
                        addInstance = true;
                    }

                    if (class_name == "childRow")
                    {
                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id"), decode));
                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id", true), decode));
                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                        reportItemComments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                        num_id++;
                        addInstance = true;
                    }

                    if (class_name == "childRowLast")
                    {
                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id"), decode));
                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id", true), decode));
                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                        reportItemComments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                        num_id++;
                        resetgroupid = true;
                        addInstance = true;
                    }

                    if (class_name == "contentRow")
                    {
                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id"), decode));
                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Объект Id", true), decode));
                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                        reportItemComments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                        num_id++;
                        addInstance = true;
                    }

                    if (name == string.Empty)
                        continue;

                    try
                    {
                        FileInfo img = new FileInfo(image);
                        if (!img.Exists)
                        {
                            Print(string.Format("Элемент «{0}» не будет добавлен в отчет! - Изображение не найдено! ({1})", name, image), MessageType.Error);
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                        Print(string.Format("Элемент «{0}» не будет добавлен в отчет! - Изображение не найдено! ({1})", name, image), MessageType.Error);
                        continue;
                    }

                    if (addInstance)
                    {
                        if (isGroupInstance)
                        {
                            reportInstances.Add(new ReportItem(
                                num_id,
                                repGroup.Id,
                                reportId,
                                name,
                                image,
                                KPItemStatus.Opened,
                                reportParentGroupComments,
                                reportItemComments));
                        }
                        else
                        {
                            reportInstances.Add(new ReportItem(
                                num_id,
                                repGroup.Id,
                                reportId,
                                name,
                                image,
                                KPItemStatus.Opened,
                                reportParentGroupComments,
                                reportItemComments,
                                parentGroupId,
                                id1,
                                id2,
                                name1,
                                name2,
                                point));
                        }
                    }

                    if (resetgroupid)
                    {
                        parentGroupId = -1;
                    }
                }

                return reportInstances;
            }

            return null;
        }





        private void OnButtonCloseReportGroup(object sender, RoutedEventArgs args)
        {
            KPTaskDialog dialog = new KPTaskDialog(
                this,
                "Закрыть группу",
                "Необходимо подтверждение",
                "Вы уверены, что хотите закрыть данную группу отчетов?",
                Core.ClashesMainCollection.KPTaskDialogIcon.Ooo,
                true,
                "После закрытия все действия с отчетами будут заморожены!");
            dialog.ShowDialog();

            if (dialog.DialogResult == Core.ClashesMainCollection.KPTaskDialogResult.Ok)
            {
                if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
                {
                    ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                    group.Status = Core.ClashesMainCollection.KPItemStatus.Closed;
                    _sqliteService_MainDB.UpdateItemStatus_ByTableAndItemId(group.Status, MainDB_Enumerator.ReportGroups, group.Id);
                    UpdateReportGroups();
                }
                else
                {
                    KPTaskDialog dialogError = new KPTaskDialog(
                        this,
                        "Закрыть группу",
                        "Ошибка",
                        "Закрыть может только сотрудник BIM-отдела",
                        Core.ClashesMainCollection.KPTaskDialogIcon.Ooo,
                        true);
                    dialogError.ShowDialog();
                }
            }
        }

























        private void OnButtonImportStatus(object sender, RoutedEventArgs args)
        {
            if (DBMainService.CurrentUserDBSubDepartment.Id != 8) { return; }

            string NormalizeTitle(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                s = s.Normalize(NormalizationForm.FormKC).ToLowerInvariant();
                var sbNorm = new System.Text.StringBuilder(s.Length);
                bool prevSpace = false;
                foreach (char ch in s)
                {
                    if (char.IsLetterOrDigit(ch)) { sbNorm.Append(ch); prevSpace = false; }
                    else if (char.IsWhiteSpace(ch)) { if (!prevSpace) { sbNorm.Append(' '); prevSpace = true; } }
                }
                return sbNorm.ToString().Trim();
            }

            string PairKey(int a, int b)
            {
                if (a > b) { var t = a; a = b; b = t; }
                return $"{a}|{b}";
            }

            bool IsGroupHeader(ReportItem r) => r != null && r.ParentGroupId == -1 && r.Element_1_Id < 0 && r.Element_2_Id < 0;
            bool IsSingle(ReportItem r) => r != null && r.ParentGroupId == -1 && r.Element_1_Id >= 0 && r.Element_2_Id >= 0;

            List<ReportItem> LoadReportItemsSafely(Report rep)
            {
                try
                {
                    var svc = new Services.SQLite.SQLiteService_ReportItemsDB(rep.PathToReportInstance);
                    var arr = svc.GetAllReporItems();
                    return arr != null ? arr.ToList() : new List<ReportItem>();
                }
                catch { return new List<ReportItem>(); }
            }

            Dictionary<string, List<Report>> BuildReportNameDict(IEnumerable<Report> reports)
            {
                var dict = new Dictionary<string, List<Report>>(StringComparer.Ordinal);
                foreach (var r in reports ?? Enumerable.Empty<Report>())
                {
                    if (r == null || string.IsNullOrWhiteSpace(r.Name)) continue;
                    string key = NormalizeTitle(r.Name);
                    if (!dict.TryGetValue(key, out var list)) { list = new List<Report>(); dict[key] = list; }
                    list.Add(r);
                }
                return dict;
            }

            List<ReportItemComment> GetCommentsAsList(ReportItem ri)
            {
                if (ri == null) return new List<ReportItemComment>();
                _ = ri.Comments;
                return (ri.CommentCollection ?? new ObservableCollection<ReportItemComment>()).ToList();
            }

            List<string> ToStringList(IEnumerable<ReportItemComment> comments) =>
                (comments ?? Enumerable.Empty<ReportItemComment>()).Select(c => c?.ToString() ?? string.Empty).ToList();

            // Разбор даты из закодированной строки user~dept~dd.MM.yyyy HH:mm:ss~dept2~text
            DateTime? TryParseCommentDate(string encoded)
            {
                if (string.IsNullOrWhiteSpace(encoded)) return null;
                var parts = encoded.Split('~');
                if (parts.Length < 3) return null;
                if (DateTime.TryParseExact(parts[2].Trim(), "dd.MM.yyyy HH:mm:ss", new CultureInfo("ru-RU"), DateTimeStyles.None, out var dt))
                    return dt;
                return null;
            }

            // Сортировка по дате убывания (новые сверху);
            // Строки без даты — вниз
            List<string> SortCommentsDescByDate(IEnumerable<string> all)
            {
                var withIdx = (all ?? Enumerable.Empty<string>()).Select((s, i) => new { s, i, dt = TryParseCommentDate(s) }).ToList();
                return withIdx.OrderByDescending(x => x.dt ?? DateTime.MinValue)
                    .ThenBy(x => x.dt.HasValue ? 0 : 1).ThenBy(x => x.i).Select(x => x.s).ToList();
            }

            List<string> TakeTopNByVisualOrder(IEnumerable<string> all, int n)
            {
                if (n <= 0) return new List<string>();
                return (all ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Take(n).ToList();
            }

            bool HasEligibleCommentsDst(ReportItem dstItem)
            {
                var list = ToStringList(GetCommentsAsList(dstItem));
                if (list.Count == 0) return false;
                if (!list.Any(s => s.IndexOf("<Делегирована отделу", StringComparison.OrdinalIgnoreCase) >= 0)) return false;
                var top = list.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? string.Empty;
                if (top.IndexOf("Статус изменен: <Устранено>", StringComparison.OrdinalIgnoreCase) >= 0) return false;
                return true;
            }

            // Преобразование одной закодированной строки:
            // ДАТА  = NOW, а ТЕКСТ = "ИМПОРТ [<origDate>]: <origText>"
            string TransformEncodedForImport(string encoded, DateTime now)
            {
                if (string.IsNullOrWhiteSpace(encoded)) return string.Empty;
                var parts = encoded.Split('~');
                if (parts.Length < 5) return encoded; 

                string origDate = parts[2];
                string origText = parts[4];

                parts[2] = now.ToString("dd.MM.yyyy HH:mm:ss");
                parts[4] = $"ПРЕДЫДУЩАЯ ИТЕРАЦИЯ [{origDate}]: {origText}"; 

                return string.Join("~", parts);
            }

            string PairView(int a, int b) => $"({a},{b})";

            // ВЫБОР ГРУПП И КОЛ-ВА КОММЕНТАРИЕВ
            var btn = sender as System.Windows.Controls.Button;
            ReportGroup sourceGroup = btn != null ? btn.DataContext as ReportGroup : null;
            if (sourceGroup == null) return;

            var allGroups = _sqliteService_MainDB.GetReportGroups_ByDBProject(_project).OrderBy(gr => gr.Status != KPItemStatus.Closed).ThenBy(gr => gr.Id).ToList();

            var picker = new ReportGroupPickerForm(allGroups, sourceGroup.Id);
            bool? pickRes = picker.ShowDialog();
            if (pickRes != true || picker.SelectedGroup == null) return;

            ReportGroup targetGroup = picker.SelectedGroup;
            int userNumber = picker.SelectedNumber;

            var srcReports = _sqliteService_MainDB.GetReports_ByReportGroupId(sourceGroup.Id) ?? new ObservableCollection<Report>();
            var dstReports = _sqliteService_MainDB.GetReports_ByReportGroupId(targetGroup.Id) ?? new ObservableCollection<Report>();
            var srcByName = BuildReportNameDict(srcReports);
            var dstByName = BuildReportNameDict(dstReports);
            var commonReportKeys = srcByName.Keys.Intersect(dstByName.Keys, StringComparer.Ordinal).OrderBy(s => s, StringComparer.Ordinal).ToList();

            var writePlan = new Dictionary<Report, List<(ReportItem item, List<string> transformedLines)>>();

            var reportBuilder = new System.Text.StringBuilder();
            reportBuilder.AppendLine($"SOURCE ReportGroup: {sourceGroup.GroupName}");
            reportBuilder.AppendLine($"TARGET ReportGroup: {targetGroup.GroupName}");
            reportBuilder.AppendLine($"Берём верхние из DST: {userNumber} шт.");
            reportBuilder.AppendLine("Текст переносимых комментариев преобразуется в: \"ИМПОРТ [<ориг.дата>]: <ориг.текст>\", дата становится текущей.");
            reportBuilder.AppendLine(new string('-', 100));
            reportBuilder.AppendLine();

            if (commonReportKeys.Count == 0)
            {
                reportBuilder.AppendLine("Совпадающих отчётов не найдено.");
            }
            else
            {
                foreach (var repKey in commonReportKeys)
                {
                    var srcRep = srcByName[repKey].FirstOrDefault();
                    var dstRep = dstByName[repKey].FirstOrDefault();
                    if (srcRep == null || dstRep == null) continue;

                    var srcItems = LoadReportItemsSafely(srcRep);
                    var dstItems = LoadReportItemsSafely(dstRep);

                    // ОДИНОЧНЫЕ: матч по имени + паре ID
                    var srcSingles = srcItems.Where(IsSingle).ToList();
                    var dstSingles = dstItems.Where(IsSingle).ToList();
                    var dstSinglesIndex = new Dictionary<(string name, string pair), ReportItem>();
                    foreach (var d in dstSingles)
                        dstSinglesIndex[(NormalizeTitle(d.Name), PairKey(d.Element_1_Id, d.Element_2_Id))] = d;

                    var matchedSingles = new List<(ReportItem src, ReportItem dst)>();
                    foreach (var sItem in srcSingles)
                    {
                        var key = (NormalizeTitle(sItem.Name), PairKey(sItem.Element_1_Id, sItem.Element_2_Id));
                        if (dstSinglesIndex.TryGetValue(key, out var dMatch) && HasEligibleCommentsDst(dMatch))
                            matchedSingles.Add((sItem, dMatch));
                    }

                    // ГРУППЫ: матч по имени заголовка
                    var srcHdrs = srcItems.Where(IsGroupHeader).ToList();
                    var dstHdrs = dstItems.Where(IsGroupHeader).ToList();
                    var dstHdrsByName = dstHdrs.GroupBy(h => NormalizeTitle(h.Name)).ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);

                    var matchedGroups = new List<(ReportItem srcHdr, ReportItem dstHdr)>();
                    foreach (var sHdr in srcHdrs)
                    {
                        string key = NormalizeTitle(sHdr.Name);
                        if (dstHdrsByName.TryGetValue(key, out var dHdrList))
                        {
                            var dCandidate = dHdrList.FirstOrDefault(h => HasEligibleCommentsDst(h));
                            if (dCandidate != null)
                                matchedGroups.Add((sHdr, dCandidate));
                        }
                    }

                    // Заголовок блока
                    reportBuilder.AppendLine($"Report (Navis): {srcRep.Name}");
                    if (matchedGroups.Count > 0)
                    {
                        reportBuilder.AppendLine("  Групповые:");
                        foreach (var (srcHdr, _) in matchedGroups) reportBuilder.AppendLine($"    - {srcHdr.Name}");
                    }
                    if (matchedSingles.Count > 0)
                    {
                        reportBuilder.AppendLine("  Одиночные:");
                        foreach (var (srcS, _) in matchedSingles) reportBuilder.AppendLine($"    - {srcS.Name}");
                    }
                    if (matchedGroups.Count == 0 && matchedSingles.Count == 0)
                        reportBuilder.AppendLine("  (Нет подходящих элементов по условиям комментариев)");
                    reportBuilder.AppendLine();

                    DateTime now = DateTime.Now;

                    // ГРУППЫ
                    foreach (var (srcHdr, dstHdr) in matchedGroups)
                    {
                        var srcCur = ToStringList(GetCommentsAsList(srcHdr));
                        var dstAll = ToStringList(GetCommentsAsList(dstHdr));

                        var dstTopVisual = TakeTopNByVisualOrder(dstAll, userNumber);
                        var transformed = dstTopVisual.Select(s => TransformEncodedForImport(s, now)).ToList();

                        var finalPreview = SortCommentsDescByDate(transformed.Concat(srcCur).ToList());
                        reportBuilder.AppendLine($"[GROUP] {srcHdr.Name}");
                        reportBuilder.AppendLine($"  SRC: БУДЕТ ПОСЛЕ ПЕРЕНОСА (новые сверху) — всего {finalPreview.Count}");
                        foreach (var line in finalPreview) reportBuilder.AppendLine($"    • {line}");
                        reportBuilder.AppendLine();

                        if (!writePlan.TryGetValue(srcRep, out var bucket))
                        {
                            bucket = new List<(ReportItem, List<string>)>();
                            writePlan[srcRep] = bucket;
                        }
                        bucket.Add((srcHdr, transformed));
                    }

                    // ОДИНОЧНЫЕ
                    foreach (var (srcS, dstS) in matchedSingles)
                    {
                        var srcCur = ToStringList(GetCommentsAsList(srcS));
                        var dstAll = ToStringList(GetCommentsAsList(dstS));

                        var dstTopVisual = TakeTopNByVisualOrder(dstAll, userNumber);
                        var transformed = dstTopVisual.Select(s => TransformEncodedForImport(s, DateTime.Now)).ToList();

                        var finalPreview = SortCommentsDescByDate(transformed.Concat(srcCur).ToList());
                        reportBuilder.AppendLine($"[SINGLE] {srcS.Name} {PairView(srcS.Element_1_Id, srcS.Element_2_Id)}");
                        reportBuilder.AppendLine($"  SRC: БУДЕТ ПОСЛЕ ПЕРЕНОСА (новые сверху) — всего {finalPreview.Count}");
                        foreach (var line in finalPreview) reportBuilder.AppendLine($"    • {line}");
                        reportBuilder.AppendLine();

                        if (!writePlan.TryGetValue(srcRep, out var bucket))
                        {
                            bucket = new List<(ReportItem, List<string>)>();
                            writePlan[srcRep] = bucket;
                        }
                        bucket.Add((srcS, transformed));
                    }

                    reportBuilder.AppendLine(new string('-', 100));
                    reportBuilder.AppendLine();
                }
            }

            // Отладчик
            void SavePreview(string text)
            {
                try
                {
                    var dlg = new Microsoft.Win32.SaveFileDialog
                    {
                        Title = "Сохранить отчёт по импорту комментариев",
                        Filter = "Text file (*.txt)|*.txt",
                        FileName = $"ImportComments_PREVIEW_{DateTime.Now:yyyyMMdd_HHmm}.txt"
                    };
                    if (dlg.ShowDialog() == true)
                        System.IO.File.WriteAllText(dlg.FileName, text, System.Text.Encoding.UTF8);
                }
                catch {}
            }

            // Сохранение в БД
            void ApplyWritePlan(Dictionary<Report, List<(ReportItem item, List<string> transformedLines)>> plan)
            {
                if (plan == null || plan.Count == 0) return;

                var applyDialog = new KPTaskDialog(this, "Импорт комментариев", "Подтверждение", "Применить перенос комментариев?\n",
                    KPTaskDialogIcon.Question, true, "Отката нет. При необходимости сделайте резервную копию.");
                applyDialog.ShowDialog();
                if (applyDialog.DialogResult != KPTaskDialogResult.Ok) return;

                foreach (var kv in plan)
                {
                    Report srcRep = kv.Key;
                    var changes = kv.Value;

                    var svc = new Services.SQLite.SQLiteService_ReportItemsDB(srcRep.PathToReportInstance);

                    foreach (var (srcItem, transformedLines) in changes)
                    {
                        if (transformedLines == null || transformedLines.Count == 0) continue;

                        for (int i = transformedLines.Count - 1; i >= 0; i--)
                        {
                            string encoded = transformedLines[i];
                            if (string.IsNullOrWhiteSpace(encoded)) continue;

                            svc.PrependEncodedComment_ByReportItem(encoded, srcItem, overrideDateToNow: false);

                            var existing = srcItem.CommentCollection?.Select(c => c.ToString()) ?? Enumerable.Empty<string>();
                            var updated = new[] { encoded }.Concat(existing);
                            srcItem.Comments = string.Join(ClashesMainCollection.StringSeparatorItem, updated);
                        }

                        bool enableAutoDelegation = true; // Автоделегирование
                        if (enableAutoDelegation)
                        {
                            foreach (var encoded in transformedLines)
                            {
                                string deptCode = TryExtractDeptCode(encoded);
                                int deptId = MapDeptCodeToId(deptCode);

                                if (deptId > 0)
                                {
                                    int statusId = (int)KPItemStatus.Delegated;
                                    svc.UpdateDelegationAndStatus_ByReportItem(srcItem, deptId, statusId);

                                    srcItem.DelegatedDepartmentId = deptId;
                                    srcItem.Status = KPItemStatus.Delegated;

                                    break; 
                                }
                            }
                        }
                    }
                }

                UpdateSelectedReportGroup(sourceGroup);
            }
         
            //SavePreview(reportBuilder.ToString());    // Отладчик
            ApplyWritePlan(writePlan);                // Сохранение в БД
        }


        /// <summary>
        /// Код отдела из текста комментария
        /// </summary>
        string TryExtractDeptCode(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            const string marker = "<Делегирована отделу";
            int i = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (i < 0) return null;

            int start = i + marker.Length;
            while (start < text.Length && char.IsWhiteSpace(text[start])) start++;

            var sb = new System.Text.StringBuilder();
            while (start < text.Length && text[start] != '>')
            {
                sb.Append(text[start]);
                start++;
            }
            string code = sb.ToString().Trim();

            if (string.IsNullOrEmpty(code)) return null;
            return code.ToUpperInvariant(); 
        }

        /// <summary>
        /// Маппинг кода отдела
        /// </summary>
        int MapDeptCodeToId(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return -1;
            switch (code.Trim().ToUpperInvariant())
            {
                case "АР": return 2;
                case "КР": return 3;
                case "ОВ": return 4;
                case "ВК": return 5;
                case "ЭОМ": return 6;
                case "СС": return 7;
                case "ИТП": return 20;
                case "ПТ": return 21;
                case "АВ": return 22;
                default: return -1;
            }
        }

















































        private void OnBtnAddGroup(object sender, RoutedEventArgs args)
        {
            if (DBMainService.CurrentUserDBSubDepartment.Id != 8) { return; }

            ReportManagerCreateGroupForm groupCreateForm = new ReportManagerCreateGroupForm();

            if ((bool)groupCreateForm.ShowDialog())
            {
                _sqliteService_MainDB.PostReportGroups_NewGroupByProjectAndName(_project, groupCreateForm.CurrentReportGroup);
                UpdateReportGroups();
            }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args) => UpdateReportGroups();

        private void RecEnter(object sender, System.Windows.Input.MouseEventArgs args)
        {
            SolidColorBrush color = new SolidColorBrush(Color.FromArgb(255, 0, 255, 115));
            switch (sender.GetType().Name)
            {
                case nameof(System.Windows.Shapes.Rectangle):
                    ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill = color;
                    break;
                case nameof(TextBlock):
                    ((sender as TextBlock).DataContext as Report).Fill = color;
                    break;
                case nameof(Image):
                    ((sender as Image).DataContext as Report).Fill = color;
                    break;
            }
        }

        private void RecLeave(object sender, System.Windows.Input.MouseEventArgs args)
        {
            switch (sender.GetType().Name)
            {
                case nameof(System.Windows.Shapes.Rectangle):
                    ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill = ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill_Default;
                    break;
                case nameof(TextBlock):
                    ((sender as TextBlock).DataContext as Report).Fill = ((sender as TextBlock).DataContext as Report).Fill_Default;
                    break;
                case nameof(Image):
                    ((sender as Image).DataContext as Report).Fill = ((sender as Image).DataContext as Report).Fill_Default;
                    break;
            }
        }

        private ReportGroup GetGroupById(int groupid)
        {
            foreach (object obj in FilteredRepGroupColl)
            {
                if (obj is ReportGroup group && group.Id == groupid)
                    return group;
            }

            return null;
        }

        private void OnUp(object sender, MouseButtonEventArgs args)
        {
            if (args.ChangedButton == MouseButton.Right)
            {
                if (sender.GetType() == typeof(System.Windows.Shapes.Rectangle))
                {
                    Report report = (sender as System.Windows.Shapes.Rectangle).DataContext as Report;
                    report.GetProgress();
                }

                if (sender.GetType() == typeof(TextBlock))
                {
                    Report report = (sender as TextBlock).DataContext as Report;
                    report.GetProgress();
                }

                if (sender.GetType() == typeof(Image))
                {
                    Report report = (sender as Image).DataContext as Report;
                    report.GetProgress();
                }
            }
            if (args.ChangedButton == MouseButton.Left)
            {
                try
                {
                    if (sender.GetType() == typeof(System.Windows.Shapes.Rectangle))
                    {
                        Report report = (sender as System.Windows.Shapes.Rectangle).DataContext as Report;
                        ReportForm form = new ReportForm(this, report, GetGroupById(report.ReportGroupId));
                        form.Show();
                    }
                    if (sender.GetType() == typeof(TextBlock))
                    {
                        Report report = (sender as TextBlock).DataContext as Report;
                        ReportForm form = new ReportForm(this, report, GetGroupById(report.ReportGroupId));
                        form.Show();
                    }
                    if (sender.GetType() == typeof(Image))
                    {
                        Report report = (sender as Image).DataContext as Report;
                        ReportForm form = new ReportForm(this, report, GetGroupById(report.ReportGroupId));
                        form.Show();
                    }
                }
                catch (Exception e)
                {
                    Print(sender.GetType().FullName, MessageType.Code);
                    PrintError(e);
                }
            }
        }

        private void OnRemoveGroup(object sender, RoutedEventArgs e)
        {
            KPTaskDialog dialog = new KPTaskDialog(
                this,
                "Удалить группу",
                "Необходимо подтверждение",
                "Вы уверены, что хотите удалить данную группу отчетов?",
                Core.ClashesMainCollection.KPTaskDialogIcon.Ooo,
                true,
                "После удаления данные будут безвозвратно потеряны!");
            dialog.ShowDialog();

            if (dialog.DialogResult == Core.ClashesMainCollection.KPTaskDialogResult.Ok)
            {
                if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
                {
                    ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                    _sqliteService_MainDB.DeleteReportGroupAndReportsAndReportItems_ByReportGroupId(group.Id);
                    UpdateReportGroups();
                }
                else
                {
                    KPTaskDialog dialogError = new KPTaskDialog(
                        this,
                        "Удалить группу",
                        "Ошибка",
                        "Удалить может только сотрудник BIM-отдела",
                        Core.ClashesMainCollection.KPTaskDialogIcon.Ooo,
                        true);
                    dialogError.ShowDialog();
                }
            }
        }

        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is System.Windows.Controls.TextBox tb && tb.DataContext is ReportGroup reportGroup)
                ApplySearchToReportGroup(reportGroup);
        }

        private void ApplySearchToReportGroup(ReportGroup reportGroup)
        {
            string search = reportGroup.SearchText?.ToLower() ?? string.Empty;

            foreach (Report report in reportGroup.Reports)
            {
                report.IsReportVisible = string.IsNullOrEmpty(search) || report.Name.ToLower().Contains(search);
            }
        }

        private void ShowClosedReportGroups_Checked(object sender, RoutedEventArgs e) =>
            FilteredRepGroupColl?.Refresh();


        private void ReportNameTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _reportNameTBxData = textBox.Text.ToLower();

            FilteredRepGroupColl?.Refresh();
        }

        private void ReportBitrixIdTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _reportBitrixIdTBxData = textBox.Text.ToLower();

            FilteredRepGroupColl?.Refresh();
        }

        private void BitrixTaskBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            if (button.DataContext is SubDepartmentBtn subDepartmentBtn)
            {
                if (subDepartmentBtn.Id == -1 || subDepartmentBtn.Id == 0)
                {
                    System.Windows.MessageBox.Show(
                        $"Не удалось получить ID-задачи из Bitrix. Скорее всего задачу не привязали",
                        "Открытие задачи в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                Process.Start("chrome", $"https://kpln.bitrix24.ru/company/personal/user/{DBMainService.CurrentDBUser.BitrixUserID}/tasks/task/view/{subDepartmentBtn.Id}/");
            }
        }
    }
}

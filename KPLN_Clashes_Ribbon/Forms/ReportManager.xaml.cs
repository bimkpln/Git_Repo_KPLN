using HtmlAgilityPack;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.Services;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Tools.HTMLTools;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using Image = System.Windows.Controls.Image;
using Path = System.IO.Path;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportManager.xaml
    /// </summary>
    public partial class ReportManager : Window
    {
        private readonly DBProject _project;
        private readonly SQLiteService_MainDB _sqliteService_MainDB = new SQLiteService_MainDB();

        public ReportManager(DBProject project)
        {
            _project = project ?? throw new ArgumentNullException("\n[KPLN]: Попытка передачи пустого проекта\n");

            InitializeComponent();
            UpdateGroups();

            if (CurrentDBUser.SubDepartmentId == 8)
                btnAddGroup.Visibility = Visibility.Visible;
            else
                btnAddGroup.Visibility = Visibility.Collapsed;
        }

        public void UpdateGroups()
        {
            ObservableCollection<ReportGroup> groups = _sqliteService_MainDB.GetReportGroups_ByDBProject(_project);
            if (groups != null)
            {
                foreach (ReportGroup group in groups)
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
                        
                        group.Reports.Add(report);
                    }
                }

                this.iControllGroups.ItemsSource = groups;
            }
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
                if (CurrentDBUser.SubDepartmentId == 8)
                {
                    Report report = (sender as System.Windows.Controls.Button).DataContext as Report;
                    _sqliteService_MainDB.DeleteReportAndReportItems_ByReportId(report);
                    UpdateGroups();
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
            if (CurrentDBUser.SubDepartmentId == 8)
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
                            ObservableCollection<ReportItem> reportInstances = ParseHtmlToRepInstColelction(file, group);
                            if (reportInstances != null)
                            {
                                if (reportInstances.Count != 0)
                                {
                                    FileInfo db_FI = GenerateNewPath_DBForReportInstance(group, ++repInstIndex);

                                    //Создаю БД под item и публикую в него данные
                                    SQLiteService_ReportItemsDB sqliteService_ReportItemsDB = new SQLiteService_ReportItemsDB(db_FI.FullName);
                                    Task fiWorkTask = Task.Run(() =>
                                    {
                                        sqliteService_ReportItemsDB.CreateDbFile_ByReports();
                                        sqliteService_ReportItemsDB.PostNewItems_ByReports(reportInstances);
                                    });
                                    riWorkerTasks[taskCount++] = fiWorkTask;

                                    // Публикую запись о новом репорте
                                    _sqliteService_MainDB.PostReport_NewReport_ByNameAndReportGroup(file.Name.Replace(".html", ""), group, db_FI);
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
                        UpdateGroups();
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
        private ObservableCollection<ReportItem> ParseHtmlToRepInstColelction(FileInfo file, ReportGroup repGroup)
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
                                name,
                                id1,
                                id2,
                                name1,
                                name2,
                                image,
                                point,
                                KPItemStatus.Opened,
                                reportParentGroupComments,
                                reportItemComments,
                                -1));
                        }
                        else
                        {
                            reportInstances.Add(new ReportItem(
                                num_id,
                                repGroup.Id,
                                name,
                                id1,
                                id2,
                                name1,
                                name2,
                                image,
                                point,
                                KPItemStatus.Opened,
                                reportParentGroupComments,
                                reportItemComments,
                                parentGroupId));
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
                if (CurrentDBUser.SubDepartmentId == 8)
                {
                    ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                    group.Status = Core.ClashesMainCollection.KPItemStatus.Closed;
                    _sqliteService_MainDB.UpdateItemStatus_ByTableAndItemId(group.Status, MainDB_Enumerator.ReportGroups, group.Id);
                    UpdateGroups();
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

        private void OnBtnAddGroup(object sender, RoutedEventArgs args)
        {
            if (CurrentDBUser.SubDepartmentId != 8) { return; }
            
            TextInputForm textInputForm = new TextInputForm(this, "Введите наименование отчета:");
            textInputForm.ShowDialog();
            if (textInputForm.IsConfirmed)
            {
                _sqliteService_MainDB.PostReportGroups_NewGroupByProjectAndName(_project, textInputForm.UserComment);
                UpdateGroups();
            }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            UpdateGroups();
        }

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
            foreach (ReportGroup group in this.iControllGroups.ItemsSource as ObservableCollection<ReportGroup>)
            {
                if (group.Id == groupid)
                {
                    return group;
                }
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
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.ReportGroupId).IsEnabled, this);
                        form.Show();
                    }
                    if (sender.GetType() == typeof(TextBlock))
                    {
                        Report report = (sender as TextBlock).DataContext as Report;
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.ReportGroupId).IsEnabled, this);
                        form.Show();
                    }
                    if (sender.GetType() == typeof(Image))
                    {
                        Report report = (sender as Image).DataContext as Report;
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.ReportGroupId).IsEnabled, this);
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
                if (CurrentDBUser.SubDepartmentId == 8)
                {
                    ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                    _sqliteService_MainDB.DeleteReportGroupAndReportsAndReportItems_ByReportGroupId(group.Id);
                    UpdateGroups();
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
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            string _searchName = textBox.Text.ToLower();

            System.Windows.Controls.TextBox tbOriginal = (System.Windows.Controls.TextBox)e.OriginalSource;

            if (tbOriginal.DataContext is ReportGroup reportGroup)
            {
                foreach (Report report in reportGroup.Reports)
                {
                    if (!report.Name.ToLower().Contains(_searchName))
                        report.IsReportVisible = false;
                    else
                        report.IsReportVisible = true;
                }
            }

        }
    }
}

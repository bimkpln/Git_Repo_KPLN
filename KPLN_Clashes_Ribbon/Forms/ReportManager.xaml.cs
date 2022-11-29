﻿using HtmlAgilityPack;
using KPLN_Clashes_Ribbon.Common.Reports;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Library_DataBase.Collections;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Tools.HTMLTools;
using static KPLN_Loader.Output.Output;
using Image = System.Windows.Controls.Image;
using Path = System.IO.Path;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportManager.xaml
    /// </summary>
    public partial class ReportManager : Window
    {
        private DbProject _project;

        public ReportManager(DbProject project)
        {
            _project = project;

            InitializeComponent();

            if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
            {
                btnAddGroup.Visibility = Visibility.Visible;
            }
            else
            {
                btnAddGroup.Visibility = Visibility.Collapsed;
            }

            UpdateGroups();
        }

        public DbProject Project
        {
            get { return _project; }
        }

        public void UpdateGroups()
        {
            ObservableCollection<ReportGroup> groups = ReportGroup.GetReportGroups(_project);
            
            ObservableCollection<Report> reports = Report.GetReports();
            
            foreach (ReportGroup group in groups)
            {
                foreach (Report report in reports)
                {
                    if (report.GroupId == group.Id)
                    {
                        if (group.Status != 1)
                            report.IsGroupEnabled = Visibility.Visible;
                        else
                            report.IsGroupEnabled = Visibility.Collapsed;
                        
                        group.Reports.Add(report);
                    }
                }
            }
            this.iControllGroups.ItemsSource = groups;
        }

        private void OnBtnRemoveReport(object sender, RoutedEventArgs args)
        {
            KPTaskDialog dialog = new KPTaskDialog(this, "Удалить отчет", "Необходимо подтверждение", "Вы уверены, что хотите удалить данный отчет?", Common.Collections.KPTaskDialogIcon.Question, true, "После удаления данные о статусе и комментарии будут безвозвратно потеряны!");
            dialog.ShowDialog();
            if (dialog.DialogResult == Common.Collections.KPTaskDialogResult.Ok)
            {
                if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
                {
                    try
                    {
                        DbController.RemoveReport((sender as System.Windows.Controls.Button).DataContext as Report);
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }
                    UpdateGroups();
                }
                else
                {
                    return;
                }
            }

        }

        private void OnBtnAddReport(object sender, RoutedEventArgs args)
        {
            if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
            {
                try
                {
                    ObservableCollection<ReportInstance> ReportInstances = new ObservableCollection<ReportInstance>();
                    OpenFileDialog dialog = new OpenFileDialog();
                    dialog.Filter = "html report (*.html)|*.html";
                    dialog.Title = "Выберите отчет(ы) NavisWorks в формате .html";
                    dialog.CheckFileExists = true;
                    dialog.CheckPathExists = true;
                    dialog.Multiselect = true;
                    DialogResult result = dialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file_name in dialog.FileNames)
                        {
                            FileInfo file = new FileInfo(file_name);
                            if (file.Extension == ".html")
                            {
                                ReportInstances.Clear();
                                HtmlAgilityPack.HtmlDocument htmlSnippet = new HtmlAgilityPack.HtmlDocument();
                                using (FileStream stream = file.OpenRead())
                                {
                                    htmlSnippet.Load(stream);
                                }
                                int num_id = 0;
                                List<string> headers = new List<string>();
                                List<string> out_headers = new List<string>();
                                int group_id = -1;
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
                                    ObservableCollection<ReportComment> comments = new ObservableCollection<ReportComment>();
                                    if (class_name == "headerRow" && headers.Count == 0)
                                    {
                                        bool todecode;
                                        if (IsMainHeader(node, out out_headers, out todecode))
                                        {
                                            headers = out_headers;
                                            decode = todecode;
                                        }
                                    }
                                    if (headers.Count == 0)
                                    {
                                        continue;
                                    }
                                    if (class_name == "clashGroupRow")
                                    {
                                        isGroupInstance = true;
                                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента"), decode));
                                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента", true), decode));
                                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                                        comments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                                        num_id++;
                                        group_id = num_id;
                                        addInstance = true;
                                    }
                                    if (class_name == "childRow")
                                    {
                                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента"), decode));
                                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента", true), decode));
                                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                                        comments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                                        num_id++;
                                        addInstance = true;
                                    }
                                    if (class_name == "childRowLast")
                                    {
                                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента"), decode));
                                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента", true), decode));
                                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                                        comments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                                        num_id++;
                                        resetgroupid = true;
                                        addInstance = true;
                                    }
                                    if (class_name == "contentRow")
                                    {
                                        name = HTMLTools.GetValue(node, GetRowId(headers, "Наименование конфликта"), decode);
                                        image = Path.Combine(file.DirectoryName, HTMLTools.GetImage(node, GetRowId(headers, "Изображение")));
                                        point = HTMLTools.GetValue(node, GetRowId(headers, "Точка конфликта"), decode);
                                        id1 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента"), decode));
                                        id2 = GetId(HTMLTools.GetValue(node, GetRowId(headers, "Идентификатор элемента", true), decode));
                                        name1 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь"), decode));
                                        name2 = OptimizeV(HTMLTools.GetValue(node, GetRowId(headers, "Путь", true), decode));
                                        comments = HTMLTools.TryGetComments(HTMLTools.GetValue(node, GetRowId(headers, "Комментарии"), decode));
                                        num_id++;
                                        addInstance = true;
                                    }
                                    if (name == string.Empty) { continue; }
                                    try
                                    {
                                        FileInfo img = new FileInfo(image);
                                        if (!img.Exists)
                                        {
                                            Print(string.Format("Элемент «{0}» не будет добавлен в отчет! - Изображение не найдено! ({1})", name, image), KPLN_Loader.Preferences.MessageType.Error);
                                            continue;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        Print(string.Format("Элемент «{0}» не будет добавлен в отчет! - Изображение не найдено! ({1})", name, image), KPLN_Loader.Preferences.MessageType.Error);
                                        continue;
                                    }
                                    if (addInstance)
                                    {
                                        if (isGroupInstance)
                                        {
                                            ReportInstances.Add(new ReportInstance(
                                                num_id, 
                                                name, 
                                                id1, 
                                                id2, 
                                                name1, 
                                                name2, 
                                                image, 
                                                point, 
                                                Common.Collections.Status.Opened, 
                                                -1, 
                                                comments));
                                        }
                                        else
                                        {
                                            ReportInstances.Add(new ReportInstance(
                                                num_id, 
                                                name, 
                                                id1, 
                                                id2, 
                                                name1, 
                                                name2, 
                                                image, 
                                                point, 
                                                Common.Collections.Status.Opened, 
                                                group_id, 
                                                comments));
                                        }
                                    }
                                    if (resetgroupid)
                                    {
                                        group_id = -1;
                                    }
                                }
                                ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                                if (ReportInstances.Count != 0)
                                {
                                    DbController.AddReport(file.Name.Replace(".html", ""), group, ReportInstances);
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
                                    Print(string.Format("Не удалось создать отчет на основе файла: «{0}»;\nУбедитесь, что файл является отчетом NavisWorks и содержит следующую информацию: {1}", file.Name, string.Join(", ", parts)), KPLN_Loader.Preferences.MessageType.Error);
                                }
                            }
                        }
                        UpdateGroups();
                    }
                    else
                    { }
                }
                catch (Exception e)
                {
                    PrintError(e);
                }

            }
        }

        private void OnButtonCloseReportGroup(object sender, RoutedEventArgs args)
        {
            KPTaskDialog dialog = new KPTaskDialog(this, "Закрыть группу", "Необходимо подтверждение", "Вы уверены, что хотите закрыть данную группу отчетов?", Common.Collections.KPTaskDialogIcon.Ooo, true, "После закрытия все действия с отчетами будут заморожены!");
            dialog.ShowDialog();
            if (dialog.DialogResult == Common.Collections.KPTaskDialogResult.Ok)
            {
                if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6) 
                { 
                    ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                    try
                    {
                        group.Status = 1;
                        DbController.SetGroupValue(group.Id, "Status", 1);
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }
                    UpdateGroups();
                }
            }
        }

        private void OnBtnAddGroup(object sender, RoutedEventArgs args)
        {
            if (KPLN_Loader.Preferences.User.Department.Id != 4 && KPLN_Loader.Preferences.User.Department.Id != 6) { return; }
            try
            {
                DbProject currentDbProject;

                if (_project != null)
                {
                    currentDbProject = _project;
                }
                else
                {
                    ElementPick selectedProjectForm = SelectDbProject.CreateForm();
                    bool? dialogResult = selectedProjectForm.ShowDialog();
                    if (dialogResult != false)
                    {
                        currentDbProject = selectedProjectForm.SelectedElement.Element as DbProject;
                    }
                    else
                    {
                        currentDbProject = _project;
                    }
                }

                TextInputDialog inputName = new TextInputDialog(this, "Введите наименование отчета:");
                inputName.ShowDialog();
                if (inputName.IsConfirmed())
                {
                    DbController.AddGroup(inputName.GetLastPickedValue(), currentDbProject);
                }

                UpdateGroups();
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            UpdateGroups();
        }

        private void RecEnter(object sender, System.Windows.Input.MouseEventArgs args)
        {
            SolidColorBrush color = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 255, 115));
            try
            {
                if (sender.GetType() == typeof(System.Windows.Shapes.Rectangle))
                {
                    ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill = color;
                }
                if (sender.GetType() == typeof(TextBlock))
                {
                    ((sender as TextBlock).DataContext as Report).Fill = color;
                }
                if (sender.GetType() == typeof(Image))
                {
                    ((sender as Image).DataContext as Report).Fill = color;
                }
            }
            catch (Exception)
            { }
        }

        private void RecLeave(object sender, System.Windows.Input.MouseEventArgs args)
        {
            try
            {
                if (sender.GetType() == typeof(System.Windows.Shapes.Rectangle))
                {
                    ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill = ((sender as System.Windows.Shapes.Rectangle).DataContext as Report).Fill_Default;
                }
                if (sender.GetType() == typeof(TextBlock))
                {
                    ((sender as TextBlock).DataContext as Report).Fill = ((sender as TextBlock).DataContext as Report).Fill_Default;
                }
                if (sender.GetType() == typeof(Image))
                {
                    ((sender as Image).DataContext as Report).Fill = ((sender as Image).DataContext as Report).Fill_Default;
                }
            }
            catch (Exception)
            { }
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
                try
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
                catch (Exception)
                { }
            }
            if (args.ChangedButton == MouseButton.Left)
            {
                try
                {
                    if (sender.GetType() == typeof(System.Windows.Shapes.Rectangle))
                    {
                        Report report = (sender as System.Windows.Shapes.Rectangle).DataContext as Report;
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.GroupId).IsEnabled, this);
                        form.Show();
                    }
                    if (sender.GetType() == typeof(TextBlock))
                    {
                        Report report = (sender as TextBlock).DataContext as Report;
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.GroupId).IsEnabled, this);
                        form.Show();
                    }
                    if (sender.GetType() == typeof(Image))
                    {
                        Report report = (sender as Image).DataContext as Report;
                        ReportWindow form = new ReportWindow(report, GetGroupById(report.GroupId).IsEnabled, this);
                        form.Show();
                    }
                }
                catch (Exception e)
                {
                    Print(sender.GetType().FullName, KPLN_Loader.Preferences.MessageType.Code);
                    PrintError(e);
                }
            }
        }

        private void OnRemoveGroup(object sender, RoutedEventArgs e)
        {
            KPTaskDialog dialog = new KPTaskDialog(this, "Удалить группу", "Необходимо подтверждение", "Вы уверены, что хотите удалить данную группу отчетов?", Common.Collections.KPTaskDialogIcon.Ooo, true, "После удаления данные будут безвозвратно потеряны!");
            dialog.ShowDialog();
            if (dialog.DialogResult == Common.Collections.KPTaskDialogResult.Ok)
            {
                if (KPLN_Loader.Preferences.User.Department.Id != 4 && KPLN_Loader.Preferences.User.Department.Id != 6) { return; }

                ReportGroup group = (sender as System.Windows.Controls.Button).DataContext as ReportGroup;
                foreach (Report report in Report.GetReports())
                {
                    if (report.GroupId == group.Id)
                    {
                        DbController.RemoveReport(report);
                    }

                }
                DbController.RemoveGroup(group);
                UpdateGroups();
            }
        }

        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            string _searchName = textBox.Text.ToLower();

            System.Windows.Controls.TextBox tbOriginal = (System.Windows.Controls.TextBox)e.OriginalSource;
            ReportGroup reportGroup = tbOriginal.DataContext as ReportGroup;

            if (reportGroup != null)
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

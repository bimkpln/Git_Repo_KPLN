using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Threading;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class CopyViewSoloForm : Window
    {
        private readonly UIApplication _uiapp;
        private readonly Document _mainDocument;
        private readonly Vm _vm;

        private readonly CopyTemplatesHandler _handler;
        private readonly ExternalEvent _exEvent;

        public CopyViewSoloForm(UIApplication uiapp)
        {
            InitializeComponent();

            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));

            _uiapp = uiapp;
            _mainDocument = _uiapp.ActiveUIDocument.Document;

            _handler = new CopyTemplatesHandler();
            _exEvent = ExternalEvent.Create(_handler);

            _vm = new Vm(_uiapp, _mainDocument, _exEvent, _handler);
            DataContext = _vm;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            _vm.SelectAllTemplates(true);
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            _vm.SelectAllTemplates(false);
        }

        private void ClearSearch_Click(object sender, RoutedEventArgs e)
        {
            _vm.TemplateSearchText = string.Empty;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CopyTemplates_Click(object sender, RoutedEventArgs e)
        {
            _vm.CopySelectedTemplates();
            //Close();
        }

        private sealed class CopyTemplatesHandler : IExternalEventHandler
        {
            public Action<UIApplication> Action { get; set; }

            public void Execute(UIApplication app)
            {
                try
                {
                    Action?.Invoke(app);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("KPLN. Копирование шаблонов вида", ex.ToString());
                }
                finally
                {
                    Action = null;
                }
            }

            public string GetName()
            {
                return "KPLN. Копирование шаблонов вида";
            }
        }

        private sealed class Vm : INotifyPropertyChanged
        {
            private readonly UIApplication _uiapp;
            private readonly Document _activeDoc;

            private readonly ExternalEvent _exEvent;
            private readonly CopyTemplatesHandler _handler;

            private Document _effectiveSourceDoc;

            private List<DocRef> _allDocsPool = new List<DocRef>();

            public ObservableCollection<DocItem> TargetDocs { get; private set; }
            public ObservableCollection<DocItem> SourceDocs { get; private set; }
            public ObservableCollection<TemplateItem> TemplateItems { get; private set; }

            private ICollectionView _templateItemsView;
            public ICollectionView TemplateItemsView { get { return _templateItemsView; } }

            private string _templateSearchText;
            public string TemplateSearchText
            {
                get { return _templateSearchText; }
                set
                {
                    if (Set(ref _templateSearchText, value))
                        RefreshTemplatesFilter();
                }
            }

            private DocItem _selectedTargetDoc;
            public DocItem SelectedTargetDoc
            {
                get { return _selectedTargetDoc; }
                set
                {
                    if (Set(ref _selectedTargetDoc, value))
                        RebuildSourceList();
                }
            }

            private DocItem _selectedSourceDoc;
            public DocItem SelectedSourceDoc
            {
                get { return _selectedSourceDoc; }
                set
                {
                    if (Set(ref _selectedSourceDoc, value))
                        RebuildTemplates();
                }
            }

            public string TemplatesCountText
            {
                get
                {
                    int totalAll = TemplateItems.Count;
                    int selectedAll = TemplateItems.Count(x => x.IsSelected);
                    return "Показано: " + totalAll + " (выбрано: " + selectedAll + ")";
                }
            }

            public Vm(UIApplication uiapp, Document activeDoc, ExternalEvent exEvent, CopyTemplatesHandler handler)
            {
                _uiapp = uiapp;
                _activeDoc = activeDoc;

                _exEvent = exEvent;
                _handler = handler;

                TargetDocs = new ObservableCollection<DocItem>();
                SourceDocs = new ObservableCollection<DocItem>();
                TemplateItems = new ObservableCollection<TemplateItem>();

                _templateItemsView = CollectionViewSource.GetDefaultView(TemplateItems);
                _templateItemsView.Filter = TemplateFilter;

                _templateSearchText = string.Empty;

                LoadAllDocuments();

                SelectedTargetDoc = TargetDocs.FirstOrDefault(d => DocumentEquals(d.Doc, _activeDoc))
                                  ?? TargetDocs.FirstOrDefault();

                ResetSourceDefault();
                RefreshTemplatesFilter();
            }

            private sealed class DocSwitchContext
            {
                public UIDocument PrevUiDoc;
                public Document EffectiveSourceDoc;
                public Document ActivatedDoc;
                public bool Switched;

                public Document TempOpenedDoc;
                public bool NeedCloseTempOpenedDoc;
            }

            private bool TrySwitchActiveToSource(out DocSwitchContext ctx, out string error)
            {
                ctx = new DocSwitchContext();
                error = null;

                ctx.PrevUiDoc = _uiapp.ActiveUIDocument;

                if (SelectedSourceDoc == null || SelectedSourceDoc.Doc == null)
                {
                    error = "Не выбран документ-источник.";
                    return false;
                }

                Document srcDoc = SelectedSourceDoc.Doc;

                if (srcDoc.IsLinked)
                {
                    if (string.IsNullOrWhiteSpace(srcDoc.PathName))
                    {
                        error = "Линк не имеет PathName (не сохранён/не доступен). Невозможно открыть источник.";
                        return false;
                    }

                    try
                    {
                        ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(srcDoc.PathName);
                        var openOpts = new OpenOptions();
                        Document opened = _uiapp.Application.OpenDocumentFile(modelPath, openOpts);

                        if (opened == null)
                        {
                            error = "Не удалось открыть файл линка как документ-источник.";
                            return false;
                        }

                        ctx.TempOpenedDoc = opened;
                        ctx.NeedCloseTempOpenedDoc = true;

                        ctx.EffectiveSourceDoc = opened;
                        ctx.ActivatedDoc = ctx.PrevUiDoc != null ? ctx.PrevUiDoc.Document : null;
                        ctx.Switched = false;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        error = "Не удалось открыть файл линка: " + ex.Message;
                        return false;
                    }
                }

                if (_uiapp.ActiveUIDocument != null && object.ReferenceEquals(_uiapp.ActiveUIDocument.Document, srcDoc))
                {
                    ctx.EffectiveSourceDoc = srcDoc;
                    ctx.ActivatedDoc = srcDoc;
                    ctx.Switched = false;
                    return true;
                }

                if (string.IsNullOrWhiteSpace(srcDoc.PathName))
                {
                    error = "Документ-источник не сохранён (нет PathName), активировать его невозможно. Сохраните документ и повторите.";
                    return false;
                }

                try
                {
                    UIDocument u = _uiapp.OpenAndActivateDocument(srcDoc.PathName);
                    if (u == null || u.Document == null)
                    {
                        error = "Не удалось активировать документ-источник.";
                        return false;
                    }

                    ctx.EffectiveSourceDoc = u.Document;
                    ctx.ActivatedDoc = u.Document;
                    ctx.Switched = true;
                    return true;
                }
                catch (Exception ex)
                {
                    error = "Не удалось активировать документ-источник: " + ex.Message;
                    return false;
                }
            }

            private void RestoreActiveDocument(DocSwitchContext ctx)
            {
                try
                {
                    if (ctx == null) return;

                    if (ctx.NeedCloseTempOpenedDoc && ctx.TempOpenedDoc != null)
                    {
                        try { ctx.TempOpenedDoc.Close(false); }
                        catch { }
                    }

                    if (!ctx.Switched) return;
                    if (ctx.PrevUiDoc == null) return;

                    Document prevDoc = ctx.PrevUiDoc.Document;
                    if (prevDoc == null) return;

                    if (_uiapp.ActiveUIDocument != null && object.ReferenceEquals(_uiapp.ActiveUIDocument.Document, prevDoc))
                        return;

                    if (!string.IsNullOrWhiteSpace(prevDoc.PathName))
                    {
                        _uiapp.OpenAndActivateDocument(prevDoc.PathName);
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", ex.Message);
                }
            }

            public bool CopySelectedTemplates()
            {
                if (_exEvent == null || _handler == null)
                {
                    TaskDialog.Show("Копирование шаблонов", "ExternalEvent не инициализирован.");
                    return false;
                }

                if (SelectedTargetDoc == null || SelectedSourceDoc == null)
                {
                    TaskDialog.Show("Копирование шаблонов", "Выбери документы: откуда и куда копировать.");
                    return false;
                }

                Document targetDoc = SelectedTargetDoc.Doc;
                Document sourceDoc = SelectedSourceDoc.Doc;

                if (targetDoc == null || sourceDoc == null)
                {
                    TaskDialog.Show("Копирование шаблонов", "Документы недоступны.");
                    return false;
                }

                if (string.Equals(GetDocKey(targetDoc), GetDocKey(sourceDoc), StringComparison.OrdinalIgnoreCase))
                {
                    TaskDialog.Show("Копирование шаблонов",
                        "Вы выбрали один и тот же файл в обоих списках. Выберите разные документы.");
                    return false;
                }

                var selectedNames = TemplateItems
                    .Where(x => x.IsSelected)
                    .Select(x => x.Name)
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                if (selectedNames.Count == 0)
                {
                    TaskDialog.Show("Копирование шаблонов", "Не выбрано ни одного шаблона.");
                    return false;
                }

                var targetItemSnap = SelectedTargetDoc;
                var sourceItemSnap = SelectedSourceDoc;

                _handler.Action = (uiapp) =>
                {
                    CopySelectedTemplates_Impl(uiapp, targetItemSnap, sourceItemSnap, selectedNames);
                };

                _exEvent.Raise();
                return true;
            }

            private void CopySelectedTemplates_Impl(UIApplication uiapp, DocItem targetItem, DocItem sourceItem, List<string> selectedTemplateNames)
            {
                DocSwitchContext sw = null;
                string swError = null;

                try
                {
                    Document targetDoc = targetItem != null ? targetItem.Doc : null;
                    Document sourceDocOriginal = sourceItem != null ? sourceItem.Doc : null;

                    if (targetDoc == null || sourceDocOriginal == null)
                    {
                        TaskDialog.Show("Копирование шаблонов", "Документы недоступны.");
                        return;
                    }

                    if (!TrySwitchActiveToSource(out sw, out swError))
                    {
                        TaskDialog.Show("Копирование шаблонов", swError);
                        return;
                    }

                    Document sourceDoc = sw.EffectiveSourceDoc;
                    if (sourceDoc == null)
                    {
                        TaskDialog.Show("Копирование шаблонов", "Не удалось получить документ-источник для копирования.");
                        return;
                    }

                    var sourceTemplatesByName = new FilteredElementCollector(sourceDoc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => v.IsTemplate)
                        .GroupBy(v => v.Name, StringComparer.CurrentCultureIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.CurrentCultureIgnoreCase);

                    var targetTemplatesByName = GetTargetTemplatesByName(targetDoc);

                    var copied = new List<string>();
                    var skipped = new List<string>();

                    using (var tg = new TransactionGroup(targetDoc, "KPLN. Копирование шаблонов вида"))
                    {
                        tg.Start();

                        foreach (string srcName in selectedTemplateNames)
                        {
                            View srcTemplate;
                            if (!sourceTemplatesByName.TryGetValue(srcName, out srcTemplate) || srcTemplate == null)
                            {
                                skipped.Add(srcName + " (не найден в источнике)");
                                continue;
                            }

                            ElementId srcId = srcTemplate.Id;

                            View existingTargetTemplate;
                            bool hasConflict = targetTemplatesByName.TryGetValue(srcName, out existingTargetTemplate);

                            if (!hasConflict)
                            {
                                string tempName = BuildTempName(srcName);
                                View newCopied = CopyTemplate_WithName(targetDoc, sourceDoc, srcId, tempName);

                                if (newCopied == null)
                                {
                                    skipped.Add(srcName + " (ошибка CopyElements)");
                                    continue;
                                }

                                string finalName = RenameView_ToExactOrUnique(targetDoc, newCopied.Id, srcName);
                                copied.Add(finalName);

                                var v = targetDoc.GetElement(newCopied.Id) as View;
                                if (v != null && v.IsTemplate)
                                    targetTemplatesByName[v.Name] = v;

                                continue;
                            }

                            int assignedCount = CountViewsAssignedToTemplate(targetDoc, existingTargetTemplate.Id);
                            var action = AskDuplicateAction(srcName, assignedCount);

                            if (action == DuplicateAction.CancelOrSkip)
                            {
                                skipped.Add(srcName + " (пропущено пользователем)");
                                continue;
                            }

                            if (action == DuplicateAction.CopyWithPrefix || action == DuplicateAction.CopyWithPrefixAndAssign)
                            {
                                string tempName = BuildTempName(srcName);
                                View newCopied = CopyTemplate_WithName(targetDoc, sourceDoc, srcId, tempName);

                                if (newCopied == null)
                                {
                                    skipped.Add(srcName + " (ошибка CopyElements)");
                                    continue;
                                }

                                string prefixed = BuildCopyName(srcName);
                                string finalCopyName = RenameView_ToExactOrUnique(targetDoc, newCopied.Id, prefixed);
                                copied.Add(finalCopyName);

                                if (action == DuplicateAction.CopyWithPrefixAndAssign)
                                {
                                    var viewIds = GetViewsAssignedToTemplate(targetDoc, existingTargetTemplate.Id);
                                    AssignTemplateToViews(targetDoc, viewIds, newCopied.Id);
                                }

                                var v = targetDoc.GetElement(newCopied.Id) as View;
                                if (v != null && v.IsTemplate)
                                    targetTemplatesByName[v.Name] = v;

                                continue;
                            }

                            {
                                bool assign = (action == DuplicateAction.ReplaceOldAndAssign);

                                List<ElementId> oldAssignedViews = GetViewsAssignedToTemplate(targetDoc, existingTargetTemplate.Id);

                                if (oldAssignedViews.Count > 0)
                                    UnassignTemplateFromViews(targetDoc, oldAssignedViews);

                                if (!DeleteElementSafeResult(targetDoc, existingTargetTemplate.Id))
                                {
                                    skipped.Add(srcName + " (не удалось удалить существующий шаблон)");
                                    continue;
                                }

                                string tempName = BuildTempName(srcName);
                                View newCopied = CopyTemplate_WithName(targetDoc, sourceDoc, srcId, tempName);

                                if (newCopied == null)
                                {
                                    skipped.Add(srcName + " (ошибка CopyElements после удаления старого)");
                                    continue;
                                }

                                string finalName = RenameView_ToExactOrUnique(targetDoc, newCopied.Id, srcName);

                                if (assign && oldAssignedViews.Count > 0)
                                    AssignTemplateToViews(targetDoc, oldAssignedViews, newCopied.Id);

                                copied.Add(finalName);

                                var v = targetDoc.GetElement(newCopied.Id) as View;
                                if (v != null && v.IsTemplate)
                                    targetTemplatesByName[v.Name] = v;

                                continue;
                            }
                        }

                        tg.Assimilate();
                    }

                    ShowCopyReport(sourceDoc, targetDoc, copied, skipped);
                }
                finally
                {
                    RestoreActiveDocument(sw);
                }
            }

            private static View CopyTemplate_WithName(Document targetDoc, Document sourceDoc, ElementId sourceTemplateId, string tempName)
            {
                View newTpl = CopyOneTemplate(targetDoc, sourceDoc, sourceTemplateId);
                if (newTpl == null) return null;

                RenameView_Exact(targetDoc, newTpl.Id, tempName);

                return targetDoc.GetElement(newTpl.Id) as View;
            }

            private static void RenameView_Exact(Document doc, ElementId id, string newName)
            {
                try
                {
                    using (var t = new Transaction(doc, "KPLN. Переименование temp-шаблона"))
                    {
                        t.Start();

                        var v = doc.GetElement(id) as View;
                        if (v != null)
                            v.Name = newName;

                        t.Commit();
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", ex.ToString());
                }
            }

            private static string RenameView_ToExactOrUnique(Document doc, ElementId id, string desiredName)
            {
                RenameView_Exact(doc, id, desiredName);

                var v = doc.GetElement(id) as View;
                if (v == null) return desiredName;

                if (string.Equals(v.Name, desiredName, StringComparison.CurrentCultureIgnoreCase))
                    return v.Name;

                string unique = GetUniqueViewName(doc, desiredName);
                if (!string.Equals(unique, v.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    RenameView_Exact(doc, id, unique);
                    v = doc.GetElement(id) as View;
                    if (v != null) return v.Name;
                }

                return v.Name;
            }

            private static string BuildTempName(string originalName)
            {
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                return "__TMP__" + originalName + "__" + stamp;
            }

            private static string BuildCopyName(string originalName)
            {
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                return originalName + "_" + stamp;
            }

            private static string GetUniqueViewName(Document doc, string desiredName)
            {
                string baseName = (desiredName ?? string.Empty).Trim();
                if (baseName.Length == 0) baseName = "View Template";

                var used = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
                foreach (View v in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
                    used.Add(v.Name);

                if (!used.Contains(baseName))
                    return baseName;

                for (int i = 1; i < 10000; i++)
                {
                    string candidate = baseName + " (" + i + ")";
                    if (!used.Contains(candidate))
                        return candidate;
                }

                return baseName + " (" + Guid.NewGuid().ToString("N").Substring(0, 6) + ")";
            }

            private static Dictionary<string, View> GetTargetTemplatesByName(Document targetDoc)
            {
                var dict = new Dictionary<string, View>(StringComparer.CurrentCultureIgnoreCase);

                foreach (View v in new FilteredElementCollector(targetDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(x => x.IsTemplate))
                {
                    if (!dict.ContainsKey(v.Name))
                        dict[v.Name] = v;
                }

                return dict;
            }

            private static int CountViewsAssignedToTemplate(Document doc, ElementId templateId)
            {
                int count = 0;
                foreach (View v in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
                {
                    if (v.IsTemplate) continue;
                    if (v.ViewTemplateId == templateId) count++;
                }
                return count;
            }

            private static List<ElementId> GetViewsAssignedToTemplate(Document doc, ElementId templateId)
            {
                var ids = new List<ElementId>();

                foreach (View v in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>())
                {
                    if (v.IsTemplate) continue;
                    if (v.ViewTemplateId == templateId)
                        ids.Add(v.Id);
                }

                return ids;
            }

            private static void UnassignTemplateFromViews(Document doc, IEnumerable<ElementId> viewIds)
            {
                using (var t = new Transaction(doc, "KPLN. Снять старый шаблон с видов"))
                {
                    t.Start();

                    foreach (var id in viewIds)
                    {
                        var v = doc.GetElement(id) as View;
                        if (v == null) continue;
                        if (v.IsTemplate) continue;

                        v.ViewTemplateId = ElementId.InvalidElementId;
                    }

                    t.Commit();
                }
            }

            private static void AssignTemplateToViews(Document doc, IEnumerable<ElementId> viewIds, ElementId templateId)
            {
                using (var t = new Transaction(doc, "KPLN. Назначить шаблон на виды"))
                {
                    t.Start();

                    foreach (var id in viewIds)
                    {
                        var v = doc.GetElement(id) as View;
                        if (v == null) continue;
                        if (v.IsTemplate) continue;

                        v.ViewTemplateId = templateId;
                    }

                    t.Commit();
                }
            }

            private static bool DeleteElementSafeResult(Document doc, ElementId id)
            {
                try
                {
                    using (var t = new Transaction(doc, "KPLN. Удаление элемента"))
                    {
                        t.Start();
                        doc.Delete(id);
                        t.Commit();
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", ex.ToString());
                    return false;
                }
            }

            private static DuplicateAction AskDuplicateAction(string templateName, int assignedViewsCount)
            {
                string msg = "Уже есть шаблон с таким именем:\n"
                           + templateName
                           + "\n\n"
                           + "Назначен на видах: " + assignedViewsCount
                           + "\n\n"
                           + "Выберите действие:";

                var td = new TaskDialog("Конфликт имени шаблона");
                td.MainInstruction = "Найден шаблон с таким же именем";
                td.MainContent = msg;

                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                    "1) Копировать с новым именем (ИМЯ_ДАТА)", "");

                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    "2) Копировать с новым именем (ИМЯ_ДАТА) и назначить на виды",
                    "Новый шаблон будет назначен видам, которые сейчас используют старый");

                //td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,"3) Заменить существующий шаблон новым", "Старый шаблон будет снят с видов и удалён. Новый получит это имя.");

                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink4,
                    "3) Заменить существующий шаблон новым и назначить на виды",
                    "Старый шаблон будет удалён, новый получит это имя и будет назначен на виды");

                td.CommonButtons = TaskDialogCommonButtons.Cancel;
                td.DefaultButton = TaskDialogResult.Cancel;

                TaskDialogResult res = td.Show();

                if (res == TaskDialogResult.CommandLink1) return DuplicateAction.CopyWithPrefix;
                if (res == TaskDialogResult.CommandLink2) return DuplicateAction.CopyWithPrefixAndAssign;
                //if (res == TaskDialogResult.CommandLink3) return DuplicateAction.ReplaceOldKeepName;
                if (res == TaskDialogResult.CommandLink4) return DuplicateAction.ReplaceOldAndAssign;

                return DuplicateAction.CancelOrSkip;
            }

            private enum DuplicateAction
            {
                CancelOrSkip = 0,
                CopyWithPrefix = 1,
                CopyWithPrefixAndAssign = 2,
                ReplaceOldKeepName = 3,
                ReplaceOldAndAssign = 4
            }

            private static View CopyOneTemplate(Document targetDoc, Document sourceDoc, ElementId sourceTemplateId)
            {
                try
                {
                    using (var t = new Transaction(targetDoc, "KPLN. Копирование шаблонов из документа"))
                    {
                        t.Start();

                        var ids = new List<ElementId> { sourceTemplateId };
                        var copyOpts = new CopyPasteOptions();

                        ICollection<ElementId> newIds = ElementTransformUtils.CopyElements(
                            sourceDoc,
                            ids,
                            targetDoc,
                            Transform.Identity,
                            copyOpts);

                        t.Commit();

                        if (newIds == null) return null;

                        ElementId newId = newIds.FirstOrDefault();
                        if (newId == null || newId == ElementId.InvalidElementId) return null;

                        return targetDoc.GetElement(newId) as View;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", ex.ToString());
                    return null;
                }
            }

            private void LoadAllDocuments()
            {
                _allDocsPool = BuildAllDocsPool(_uiapp.Application);

                TargetDocs.Clear();
                foreach (var d in _allDocsPool)
                    TargetDocs.Add(new DocItem(d.Doc, d.IsLink));

                RebuildSourceList();
            }

            private void RebuildSourceList()
            {
                SourceDocs.Clear();

                foreach (var d in _allDocsPool)
                    SourceDocs.Add(new DocItem(d.Doc, d.IsLink));

                bool needReset = _selectedSourceDoc == null
                    || !SourceDocs.Any(x => string.Equals(GetDocKey(x.Doc), GetDocKey(_selectedSourceDoc.Doc), StringComparison.OrdinalIgnoreCase));

                if (needReset)
                    ResetSourceDefault();

                OnPropertyChanged(nameof(TemplatesCountText));
            }

            private void ResetSourceDefault()
            {
                Document targetDoc = SelectedTargetDoc != null ? SelectedTargetDoc.Doc : null;

                DocItem second = null;

                if (targetDoc != null)
                {
                    second = SourceDocs.FirstOrDefault(x =>
                        !x.IsLink &&
                        !string.Equals(GetDocKey(x.Doc), GetDocKey(targetDoc), StringComparison.OrdinalIgnoreCase));
                }

                if (second == null)
                    second = SourceDocs.FirstOrDefault(x => !x.IsLink);

                SelectedSourceDoc = second ?? SourceDocs.FirstOrDefault();
            }

            private void RebuildTemplates()
            {
                TemplateItems.Clear();

                Document src = _effectiveSourceDoc ?? (SelectedSourceDoc != null ? SelectedSourceDoc.Doc : null);

                if (src == null)
                {
                    RefreshTemplatesFilter();
                    return;
                }

                var templates = new FilteredElementCollector(src)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.IsTemplate)
                    .OrderBy(v => v.Name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                foreach (var v in templates)
                {
                    var ti = new TemplateItem(v.Name, v.Id);
                    ti.PropertyChanged += TemplateItem_PropertyChanged;
                    TemplateItems.Add(ti);
                }

                RefreshTemplatesFilter();
            }

            private void TemplateItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
            {
                if (e != null && e.PropertyName == nameof(TemplateItem.IsSelected))
                    OnPropertyChanged(nameof(TemplatesCountText));
            }

            public void SelectAllTemplates(bool select)
            {
                foreach (var t in TemplateItems)
                    t.IsSelected = select;

                OnPropertyChanged(nameof(TemplatesCountText));
            }

            private bool TemplateFilter(object obj)
            {
                var item = obj as TemplateItem;
                if (item == null) return false;

                if (string.IsNullOrWhiteSpace(_templateSearchText))
                    return true;

                return item.Name.IndexOf(_templateSearchText, StringComparison.CurrentCultureIgnoreCase) >= 0;
            }

            private void RefreshTemplatesFilter()
            {
                Dispatcher dispatcher = null;

                if (Application.Current != null)
                    dispatcher = Application.Current.Dispatcher;

                if (dispatcher != null)
                {
                    dispatcher.BeginInvoke(
                        DispatcherPriority.Background,
                        new Action(() =>
                        {
                            if (_templateItemsView != null)
                                _templateItemsView.Refresh();

                            OnPropertyChanged(nameof(TemplatesCountText));
                        }));
                }
                else
                {
                    if (_templateItemsView != null)
                        _templateItemsView.Refresh();

                    OnPropertyChanged(nameof(TemplatesCountText));
                }
            }

            private static List<DocRef> BuildAllDocsPool(Autodesk.Revit.ApplicationServices.Application app)
            {
                var result = new List<DocRef>();
                var hostSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var hosts = app.Documents
                    .Cast<Document>()
                    .Where(d => d != null && !d.IsFamilyDocument && !d.IsLinked)
                    .OrderBy(d => d.Title, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                foreach (var host in hosts)
                {
                    string hostKey = GetDocKey(host);
                    if (!hostSeen.Add(hostKey))
                        continue;

                    result.Add(new DocRef(host, false));

                    foreach (var linkDoc in GetLinkedDocuments(host).OrderBy(d => d.Title, StringComparer.CurrentCultureIgnoreCase))
                        result.Add(new DocRef(linkDoc, true));
                }

                return result;
            }

            private static string GetDocKey(Document doc)
            {
                if (doc == null) return string.Empty;
                if (!string.IsNullOrWhiteSpace(doc.PathName))
                    return doc.PathName;
                return doc.Title;
            }

            private static IEnumerable<Document> GetLinkedDocuments(Document hostDoc)
            {
                var linkInstances = new FilteredElementCollector(hostDoc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                foreach (var inst in linkInstances)
                {
                    Document linkDoc = inst.GetLinkDocument();
                    if (linkDoc != null)
                        yield return linkDoc;
                }
            }

            private static bool DocumentEquals(Document a, Document b)
            {
                if (object.ReferenceEquals(a, b)) return true;
                if (a == null || b == null) return false;

                string aKey = GetDocKey(a);
                string bKey = GetDocKey(b);

                return string.Equals(aKey, bKey, StringComparison.OrdinalIgnoreCase);
            }

            private static void ShowCopyReport(Document sourceDoc, Document targetDoc, List<string> copiedNames, List<string> skippedNames)
            {
                var sb = new StringBuilder();

                sb.AppendLine("Откуда:");
                sb.AppendLine(sourceDoc.Title);
                sb.AppendLine();
                sb.AppendLine("Куда:");
                sb.AppendLine(targetDoc.Title);
                sb.AppendLine();

                if (copiedNames.Count > 0)
                {
                    sb.AppendLine("Скопированы/созданы:");
                    foreach (string n in copiedNames)
                        sb.AppendLine("- " + n);
                }
                else
                {
                    sb.AppendLine("Ничего не скопировано.");
                }

                if (skippedNames.Count > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("Пропущено/ошибки:");
                    foreach (string n in skippedNames)
                        sb.AppendLine("- " + n);
                }

                TaskDialog.Show("Результат копирования", sb.ToString());
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string propName = null)
            {
                var handler = PropertyChanged;
                if (handler != null)
                    handler(this, new PropertyChangedEventArgs(propName));
            }

            private bool Set<T>(ref T field, T value, [CallerMemberName] string propName = null)
            {
                if (EqualityComparer<T>.Default.Equals(field, value)) return false;
                field = value;
                OnPropertyChanged(propName);
                return true;
            }
        }

        private sealed class DocItem
        {
            public Document Doc { get; private set; }
            public string Name { get; private set; }
            public bool IsLink { get; private set; }

            public Thickness DisplayMargin
            {
                get { return IsLink ? new Thickness(2, 0, 0, 0) : new Thickness(0); }
            }

            public DocItem(Document doc, bool isLink)
            {
                Doc = doc;
                IsLink = isLink;
                Name = doc != null ? doc.Title : "<null>";
            }

            public override string ToString()
            {
                return Name;
            }
        }

        private sealed class DocRef
        {
            public Document Doc;
            public bool IsLink;

            public DocRef(Document doc, bool isLink)
            {
                Doc = doc;
                IsLink = isLink;
            }
        }

        private sealed class TemplateItem : INotifyPropertyChanged
        {
            public string Name { get; private set; }
            public ElementId Id { get; private set; }

            private bool _isSelected;

            public bool IsSelected
            {
                get { return _isSelected; }
                set
                {
                    if (_isSelected == value) return;
                    _isSelected = value;

                    var handler = PropertyChanged;
                    if (handler != null)
                        handler(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }

            public TemplateItem(string name, ElementId id)
            {
                Name = name;
                Id = id;
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }
    }
}
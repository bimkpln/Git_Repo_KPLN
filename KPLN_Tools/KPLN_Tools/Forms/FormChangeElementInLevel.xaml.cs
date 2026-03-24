using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using Grid = System.Windows.Controls.Grid;
using TextRange = System.Windows.Documents.TextRange;

namespace KPLN_Tools.Forms
{
    public partial class FormChangeElementInLevel : Window
    {
        private readonly Document _doc;
        private bool _isProcessing = false;

        /// <summary>
        /// Категории, которые участвуют в отображении и обработке.
        /// </summary>
        private readonly List<BuiltInCategory> _targetCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_IOSModelGroups,
            BuiltInCategory.OST_GenericModel
        };

        private List<Level> _levels = new List<Level>();

        public FormChangeElementInLevel(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            LoadLevels();
            InitializeProgressState();
        }

        /// <summary>
        /// Загружает уровни проекта и заполняет комбобоксы.
        /// </summary>
        private void LoadLevels()
        {
            _levels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .ToList();

            CbSourceLevel.ItemsSource = _levels;
            CbTargetLevel.ItemsSource = _levels;

            if (_levels.Count > 0)
            {
                CbSourceLevel.SelectedIndex = 0;
                CbTargetLevel.SelectedIndex = 0;
            }

            RefreshLists();
        }

        private void InitializeProgressState()
        {
            TbStatus.Text = "Готов к запуску";
            PbProgress.Minimum = 0;
            PbProgress.Maximum = 1;
            PbProgress.Value = 0;
        }

        private void CbSourceLevel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshSourceList();
        }

        private void CbTargetLevel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshTargetList();
        }

        private void RefreshLists()
        {
            RefreshSourceList();
            RefreshTargetList();
        }

        private void RefreshSourceList()
        {
            Level selectedLevel = CbSourceLevel.SelectedItem as Level;
            if (selectedLevel == null)
            {
                LbSourceElements.ItemsSource = null;
                if (TbSourceSummary != null)
                    TbSourceSummary.Text = "Нет данных";
                return;
            }

            LevelElementsInfo info = GetLevelElementsInfo(selectedLevel);
            LbSourceElements.ItemsSource = info.Items;

            if (TbSourceSummary != null)
            {
                TbSourceSummary.Text =
                    "Групп: " + info.GroupCount +
                    " | Элементов: " + info.StandaloneElementCount +
                    " | Всего: " + info.ElementCount;
            }
        }

        private void RefreshTargetList()
        {
            Level selectedLevel = CbTargetLevel.SelectedItem as Level;
            if (selectedLevel == null)
            {
                LbTargetElements.ItemsSource = null;
                if (TbTargetSummary != null)
                    TbTargetSummary.Text = "Нет данных";
                return;
            }

            LevelElementsInfo info = GetLevelElementsInfo(selectedLevel);
            LbTargetElements.ItemsSource = info.Items;

            if (TbTargetSummary != null)
            {
                TbTargetSummary.Text =
                    "Групп: " + info.GroupCount +
                    " | Элементов: " + info.StandaloneElementCount +
                    " | Всего: " + info.ElementCount;
            }
        }

        private void BtnMove_Click(object sender, RoutedEventArgs e)
        {
            Level sourceLevel = CbSourceLevel.SelectedItem as Level;
            Level targetLevel = CbTargetLevel.SelectedItem as Level;

            if (sourceLevel == null || targetLevel == null)
            {
                MessageBox.Show("Нужно выбрать оба уровня.");
                return;
            }

            if (sourceLevel.Id == targetLevel.Id)
            {
                MessageBox.Show("Исходный и целевой уровни совпадают.");
                return;
            }

            List<ElementId> idsToMove = CollectElementIdsForLevel(sourceLevel);

            if (idsToMove.Count == 0)
            {
                MessageBox.Show("На исходном уровне не найдено объектов для переноса.");
                return;
            }

            _isProcessing = true;
            SetUiEnabled(false);
            UpdateProgress(0, 1, "Подготовка...");

            LevelTransferReport report = null;

            try
            {
                ElementLevelTransferService service = new ElementLevelTransferService();

                Progress<string> progress = new Progress<string>(value =>
                {
                    int current;
                    int total;
                    string text;

                    if (TryParseProgress(value, out current, out total, out text))
                    {
                        UpdateProgress(current, total, text);
                    }
                    else
                    {
                        UpdateProgress(
                            PbProgress.Value >= 0 ? (int)PbProgress.Value : 0,
                            PbProgress.Maximum >= 1 ? (int)PbProgress.Maximum : 1,
                            value ?? "Обработка...");
                    }
                });

                report = service.Execute(_doc, idsToMove, targetLevel.Id, progress);
            }
            catch (Exception ex)
            {
                report = report ?? new LevelTransferReport();
                report.AddGlobalError("Глобальная ошибка: " + ExceptionTextHelper.Build(ex));
            }
            finally
            {
                RefreshLists();
                SetUiEnabled(true);
                UpdateProgress(1, 1, "Готово");
                _isProcessing = false;
            }

            ReportWindow reportWindow = new ReportWindow(report, sourceLevel, targetLevel);
            reportWindow.Owner = this;
            reportWindow.ShowDialog();
        }

        private bool TryParseProgress(string value, out int current, out int total, out string text)
        {
            current = 0;
            total = 1;
            text = value ?? string.Empty;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            string[] parts = value.Split(new[] { "||" }, StringSplitOptions.None);
            if (parts.Length != 3)
                return false;

            if (!int.TryParse(parts[0], out current))
                return false;

            if (!int.TryParse(parts[1], out total))
                return false;

            text = parts[2] ?? string.Empty;
            return true;
        }

        /// <summary>
        /// Собирает только:
        /// 1) группы как самостоятельные объекты
        /// 2) элементы, которые НЕ входят в группы
        /// Элементы внутри групп полностью исключаются.
        /// </summary>
        private List<ElementId> CollectElementIdsForLevel(Level sourceLevel)
        {
            HashSet<ElementId> result = new HashSet<ElementId>(new ElementIdComparer());

            IList<Element> all = GetAllTargetElements();

            foreach (Element element in all)
            {
                if (element == null || element.Category == null)
                    continue;

                if (!IsElementOnLevel(element, sourceLevel))
                    continue;

                BuiltInCategory bic;
                try
                {
                    bic = GetBuiltInCategory(element.Category);
                }
                catch
                {
                    continue;
                }

                // Группа — переносим как группу
                Group directGroup = element as Group;
                if (bic == BuiltInCategory.OST_IOSModelGroups && directGroup != null)
                {
                    result.Add(directGroup.Id);
                    continue;
                }

                // Элементы внутри групп полностью игнорируем
                if (element.GroupId != ElementId.InvalidElementId)
                    continue;

                // Обычный самостоятельный элемент
                result.Add(element.Id);
            }

            return result
                .OrderBy(x => x.IntegerValue)
                .ToList();
        }

        private bool IsElementOnLevel(Element element, Level level)
        {
            if (element == null || level == null)
                return false;

            BuiltInCategory bic;
            try
            {
                bic = GetBuiltInCategory(element.Category);
            }
            catch
            {
                return false;
            }

            if (bic == BuiltInCategory.OST_IOSModelGroups)
            {
                Group group = element as Group;
                if (group == null)
                    return false;

                Parameter p = GetGroupBaseLevelParameter(group);
                if (p != null && p.StorageType == StorageType.ElementId)
                {
                    ElementId id = p.AsElementId();
                    if (id != ElementId.InvalidElementId && id == level.Id)
                        return true;
                }

                if (group.LevelId != ElementId.InvalidElementId && group.LevelId == level.Id)
                    return true;

                return false;
            }

            ElementId levelId;
            if (!TryGetAssociatedLevelId(element, out levelId))
                return false;

            return levelId != ElementId.InvalidElementId && levelId == level.Id;
        }

        private bool TryGetAssociatedLevelId(Element element, out ElementId levelId)
        {
            levelId = ElementId.InvalidElementId;

            if (element == null || element.Category == null)
                return false;

            BuiltInCategory bic;
            try
            {
                bic = GetBuiltInCategory(element.Category);
            }
            catch
            {
                return false;
            }

            if (bic == BuiltInCategory.OST_PipeCurves)
            {
                MEPCurve pipe = element as MEPCurve;
                if (pipe != null && pipe.ReferenceLevel != null)
                {
                    levelId = pipe.ReferenceLevel.Id;
                    return true;
                }

                Parameter startLevelParam = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                if (startLevelParam != null && startLevelParam.StorageType == StorageType.ElementId)
                {
                    levelId = startLevelParam.AsElementId();
                    if (levelId != ElementId.InvalidElementId)
                        return true;
                }

                if (element.LevelId != ElementId.InvalidElementId)
                {
                    levelId = element.LevelId;
                    return true;
                }

                return false;
            }

            if (bic == BuiltInCategory.OST_IOSModelGroups)
            {
                Group group = element as Group;
                if (group != null)
                {
                    Parameter groupLevelParam = GetGroupBaseLevelParameter(group);
                    if (groupLevelParam != null && groupLevelParam.StorageType == StorageType.ElementId)
                    {
                        levelId = groupLevelParam.AsElementId();
                        if (levelId != ElementId.InvalidElementId)
                            return true;
                    }

                    if (group.LevelId != ElementId.InvalidElementId)
                    {
                        levelId = group.LevelId;
                        return true;
                    }
                }

                return false;
            }

            Parameter levelParam = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            if (levelParam == null)
                levelParam = element.get_Parameter(BuiltInParameter.LEVEL_PARAM);

            if (levelParam != null && levelParam.StorageType == StorageType.ElementId)
            {
                levelId = levelParam.AsElementId();
                if (levelId != ElementId.InvalidElementId)
                    return true;
            }

            if (element.LevelId != ElementId.InvalidElementId)
            {
                levelId = element.LevelId;
                return true;
            }

            return false;
        }

        private Parameter GetGroupBaseLevelParameter(Group group)
        {
            if (group == null)
                return null;

            Parameter p = null;

            try
            {
                p = group.LookupParameter("Базовый уровень");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.get_Parameter(BuiltInParameter.GROUP_LEVEL);
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Уровень");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Level");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            return null;
        }

        /// <summary>
        /// В списке показываем только:
        /// - группы
        /// - элементы вне групп
        /// Элементы внутри групп не показываем.
        /// </summary>
        private LevelElementsInfo GetLevelElementsInfo(Level level)
        {
            List<ElementDisplayItem> result = new List<ElementDisplayItem>();
            int groupCount = 0;
            int standaloneCount = 0;

            IList<Element> all = GetAllTargetElements();

            foreach (Element element in all)
            {
                if (element == null || element.Category == null)
                    continue;

                if (!IsElementOnLevel(element, level))
                    continue;

                BuiltInCategory bic;
                try
                {
                    bic = GetBuiltInCategory(element.Category);
                }
                catch
                {
                    continue;
                }

                // Показываем группы отдельно
                Group directGroup = element as Group;
                if (bic == BuiltInCategory.OST_IOSModelGroups && directGroup != null)
                {
                    result.Add(new ElementDisplayItem(element, _doc));
                    groupCount++;
                    continue;
                }

                // Элементы внутри групп не показываем вообще
                if (element.GroupId != ElementId.InvalidElementId)
                    continue;

                // Показываем только самостоятельные элементы
                result.Add(new ElementDisplayItem(element, _doc));
                standaloneCount++;
            }

            result = result
                .OrderBy(x => x.SortPriority)
                .ThenBy(x => x.Name)
                .ThenBy(x => x.Id)
                .ToList();

            return new LevelElementsInfo(result, groupCount, standaloneCount);
        }

        private IList<Element> GetAllTargetElements()
        {
            List<ElementId> categoryIds = new List<ElementId>();
            foreach (BuiltInCategory bic in _targetCategories)
                categoryIds.Add(new ElementId(bic));

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categoryIds);

            return new FilteredElementCollector(_doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();
        }

        private void SetUiEnabled(bool enabled)
        {
            CbSourceLevel.IsEnabled = enabled;
            CbTargetLevel.IsEnabled = enabled;
            BtnMove.IsEnabled = enabled;
        }

        private void UpdateProgress(int current, int total, string text)
        {
            TbStatus.Text = text;
            PbProgress.Minimum = 0;
            PbProgress.Maximum = total <= 0 ? 1 : total;
            PbProgress.Value = current <= PbProgress.Maximum ? current : PbProgress.Maximum;

            DoEvents();
        }

        private static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new DispatcherOperationCallback(delegate (object parameter)
                {
                    frame.Continue = false;
                    return null;
                }),
                null);

            Dispatcher.PushFrame(frame);
        }

        private static BuiltInCategory GetBuiltInCategory(Category category)
        {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
            return (BuiltInCategory)category.Id.IntegerValue;
#else
            return category.BuiltInCategory;
#endif
        }
    }

    internal sealed class ElementLevelTransferService
    {
        private const double MoveTolerance = 1e-7;

        private Document _docCache;

        public LevelTransferReport Execute(
            Document doc,
            IList<ElementId> selectedIds,
            ElementId targetLevelId,
            IProgress<string> progress = null)
        {
            using (new DocScope(this, doc))
            {
                return ChangeLevelKeepingWorldPosition(doc, selectedIds, targetLevelId, progress);
            }
        }

        public LevelTransferReport ChangeLevelKeepingWorldPosition(
            Document doc,
            IList<ElementId> selectedIds,
            ElementId targetLevelId,
            IProgress<string> progress = null)
        {
            LevelTransferReport report = new LevelTransferReport();

            if (doc == null)
            {
                report.AddGlobalError("Документ не найден.");
                return report;
            }

            Level targetLevel = doc.GetElement(targetLevelId) as Level;
            if (targetLevel == null)
            {
                report.AddGlobalError("Целевой уровень не найден.");
                return report;
            }

            List<ElementId> normalizedIds = NormalizeSelection(doc, selectedIds);
            if (normalizedIds.Count == 0)
            {
                report.AddGlobalError("Нет элементов для обработки.");
                return report;
            }

            List<Group> groups = new List<Group>();
            List<Element> mep = new List<Element>();
            List<Element> singles = new List<Element>();

            foreach (ElementId id in normalizedIds)
            {
                Element e = doc.GetElement(id);
                if (e == null)
                    continue;

                if (e is Group)
                {
                    groups.Add((Group)e);
                    continue;
                }

                if (IsSupportedMepElement(e))
                {
                    mep.Add(e);
                    continue;
                }

                if (IsSupportedSingleElement(e))
                {
                    singles.Add(e);
                    continue;
                }

                TransferReportItem unsupported = report.GetOrCreateMainItem(e, doc);
                unsupported.ClearMessages();
                unsupported.AddError("Категория элемента не поддерживается для переноса.");
            }

            List<List<Element>> components = BuildMepConnectedComponents(mep);

            int totalOperations = groups.Count + components.Count + singles.Count;
            if (totalOperations <= 0)
                totalOperations = 1;

            int currentOperation = 0;

            foreach (Group g in groups)
            {
                currentOperation++;
                ReportProgress(progress, currentOperation, totalOperations, "Обработка группы ID: " + g.Id.IntegerValue);
                ProcessGroup(doc, g, targetLevel, report);
            }

            foreach (List<Element> component in components)
            {
                if (component.Count == 0)
                    continue;

                currentOperation++;
                ReportProgress(progress, currentOperation, totalOperations, "Обработка MEP-компонента ID: " + component[0].Id.IntegerValue);
                ProcessMepComponent(doc, component, targetLevel, report);
            }

            foreach (Element e in singles)
            {
                currentOperation++;
                ReportProgress(progress, currentOperation, totalOperations, "Обработка элемента ID: " + e.Id.IntegerValue);
                ProcessSingleElement(doc, e, targetLevel, report);
            }

            return report;
        }

        private void ReportProgress(IProgress<string> progress, int current, int total, string text)
        {
            if (progress == null)
                return;

            progress.Report(
                current.ToString(CultureInfo.InvariantCulture) + "||" +
                total.ToString(CultureInfo.InvariantCulture) + "||" +
                text);
        }

        /// <summary>
        /// Нормализация выбора:
        /// - элементы внутри групп игнорируются
        /// - группы остаются группами
        /// - обычные элементы остаются элементами
        /// </summary>
        private List<ElementId> NormalizeSelection(Document doc, IList<ElementId> selectedIds)
        {
            HashSet<ElementId> result = new HashSet<ElementId>(new ElementIdComparer());

            if (selectedIds == null)
                return result.ToList();

            foreach (ElementId id in selectedIds)
            {
                if (id == null || id == ElementId.InvalidElementId)
                    continue;

                Element e = doc.GetElement(id);
                if (e == null)
                    continue;

                // Если это элемент внутри группы — игнорируем полностью
                if (!(e is Group) && e.GroupId != ElementId.InvalidElementId)
                    continue;

                result.Add(e.Id);
            }

            return result
                .OrderBy(x => x.IntegerValue)
                .ToList();
        }

        private bool IsSupportedSingleElement(Element e)
        {
            if (e == null || e.Category == null)
                return false;

            BuiltInCategory bic;
            if (!TryGetBuiltInCategory(e.Category, out bic))
                return false;

            return bic == BuiltInCategory.OST_PlumbingFixtures
                || bic == BuiltInCategory.OST_MechanicalEquipment
                || bic == BuiltInCategory.OST_Sprinklers
                || bic == BuiltInCategory.OST_GenericModel;
        }

        private bool IsSupportedMepElement(Element e)
        {
            if (e == null || e.Category == null)
                return false;

            BuiltInCategory bic;
            if (!TryGetBuiltInCategory(e.Category, out bic))
                return false;

            return bic == BuiltInCategory.OST_PipeCurves
                || bic == BuiltInCategory.OST_PipeFitting
                || bic == BuiltInCategory.OST_PipeAccessory;
        }

        /// <summary>
        /// Перенос группы ТОЛЬКО через:
        /// - Базовый уровень
        /// - Смещение начала от уровня
        /// Группа должна визуально остаться на месте.
        /// </summary>
        private void ProcessGroup(Document doc, Group group, Level targetLevel, LevelTransferReport report)
        {
            TransferReportItem rootItem = report.GetOrCreateMainItem(group, doc);
            rootItem.ClearMessages();

            using (Transaction tx = new Transaction(doc, "Перенос группы " + group.Id.IntegerValue))
            {
                TransferFailuresPreprocessor pre = new TransferFailuresPreprocessor(rootItem);
                tx.Start();

                FailureHandlingOptions fho = tx.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(pre);
                fho.SetClearAfterRollback(true);
                tx.SetFailureHandlingOptions(fho);

                try
                {
                    Parameter levelParam = GetGroupBaseLevelParameter(group);
                    Parameter offsetParam = GetGroupBaseOffsetParameter(group);

                    if (levelParam == null)
                        throw new InvalidOperationException("Не найден параметр группы \"Базовый уровень\".");

                    if (levelParam.IsReadOnly)
                        throw new InvalidOperationException("Параметр группы \"Базовый уровень\" доступен только для чтения.");

                    if (offsetParam == null)
                        throw new InvalidOperationException("Не найден параметр группы \"Смещение начала от уровня\".");

                    if (offsetParam.IsReadOnly)
                        throw new InvalidOperationException("Параметр группы \"Смещение начала от уровня\" доступен только для чтения.");

                    Level oldLevel = GetCurrentGroupLevel(doc, group);
                    if (oldLevel == null)
                        throw new InvalidOperationException("Не удалось определить текущий базовый уровень группы.");

                    double oldLevelElevation = oldLevel.Elevation;
                    double oldOffset = offsetParam.AsDouble();
                    double newOffset = oldOffset + oldLevelElevation - targetLevel.Elevation;

                    bool restorePinned = false;
                    try
                    {
                        if (group.Pinned)
                        {
                            group.Pinned = false;
                            restorePinned = true;
                        }
                    }
                    catch
                    {
                    }

                    levelParam.Set(targetLevel.Id);
                    offsetParam.Set(newOffset);

                    doc.Regenerate();

                    string rootValidationError;
                    if (!ValidateGroupMovedToLevel(doc, group, targetLevel, out rootValidationError))
                    {
                        rootItem.AddError(rootValidationError);
                    }

                    if (restorePinned)
                    {
                        try
                        {
                            group.Pinned = true;
                        }
                        catch
                        {
                        }
                    }

                    if (rootItem.HasErrors)
                    {
                        tx.RollBack();
                    }
                    else
                    {
                        TransactionStatus status = tx.Commit();
                        if (status != TransactionStatus.Committed)
                        {
                            rootItem.AddError("Commit вернул статус " + status + ".");
                        }
                        else
                        {
                            rootItem.MarkSuccess("успешно.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tx.HasStarted())
                        tx.RollBack();

                    rootItem.AddError(ExceptionTextHelper.Build(ex));
                }
            }

            rootItem.SetChildren(new List<TransferReportItem>());
        }

        private void ProcessSingleElement(Document doc, Element e, Level targetLevel, LevelTransferReport report)
        {
            TransferReportItem item = report.GetOrCreateMainItem(e, doc);
            item.ClearMessages();

            using (Transaction tx = new Transaction(doc, "Перенос элемента " + e.Id.IntegerValue))
            {
                TransferFailuresPreprocessor pre = new TransferFailuresPreprocessor(item);
                tx.Start();

                FailureHandlingOptions fho = tx.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(pre);
                fho.SetClearAfterRollback(true);
                tx.SetFailureHandlingOptions(fho);

                try
                {
                    ElementPoseSnapshot beforePose = ElementPoseSnapshot.Create(e);
                    Level oldLevel = GetCurrentLevel(doc, e);
                    double oldLevelElevation = oldLevel != null ? oldLevel.Elevation : 0.0;
                    double oldOffset = GetOffsetValue(e);
                    double newOffset = oldOffset + oldLevelElevation - targetLevel.Elevation;

                    bool restorePinned = false;
                    try
                    {
                        if (e.Pinned)
                        {
                            e.Pinned = false;
                            restorePinned = true;
                        }
                    }
                    catch
                    {
                    }

                    SetLevelAndOffset(e, targetLevel, newOffset);

                    doc.Regenerate();
                    RestoreElementWorldPose(doc, e, beforePose);
                    doc.Regenerate();

                    string validationError;
                    if (!ValidateElementMovedToLevel(doc, e, targetLevel, out validationError))
                    {
                        item.AddError(validationError);
                    }

                    if (restorePinned)
                    {
                        try
                        {
                            e.Pinned = true;
                        }
                        catch
                        {
                        }
                    }

                    if (item.HasErrors)
                    {
                        tx.RollBack();
                    }
                    else
                    {
                        TransactionStatus status = tx.Commit();
                        if (status != TransactionStatus.Committed)
                        {
                            item.AddError("Commit вернул статус " + status + ".");
                        }
                        else
                        {
                            item.MarkSuccess("успешно.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tx.HasStarted())
                        tx.RollBack();

                    item.AddError(ExceptionTextHelper.Build(ex));
                }
            }
        }

        private void ProcessMepComponent(Document doc, List<Element> component, Level targetLevel, LevelTransferReport report)
        {
            if (component == null || component.Count == 0)
                return;

            List<TransferReportItem> items = new List<TransferReportItem>();
            foreach (Element e in component)
            {
                TransferReportItem item = report.GetOrCreateMainItem(e, doc);
                item.ClearMessages();
                items.Add(item);
            }

            int rootId = component[0].Id.IntegerValue;

            using (Transaction tx = new Transaction(doc, "Перенос MEP-компонента " + rootId))
            {
                TransferFailuresPreprocessor pre = new TransferFailuresPreprocessor(items[0]);
                tx.Start();

                FailureHandlingOptions fho = tx.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(pre);
                fho.SetClearAfterRollback(true);
                tx.SetFailureHandlingOptions(fho);

                try
                {
                    Dictionary<ElementId, ElementPoseSnapshot> poses = new Dictionary<ElementId, ElementPoseSnapshot>(new ElementIdComparer());
                    Dictionary<ElementId, bool> pinnedStates = new Dictionary<ElementId, bool>(new ElementIdComparer());

                    foreach (Element e in component)
                    {
                        poses[e.Id] = ElementPoseSnapshot.Create(e);

                        bool wasPinned = false;
                        try
                        {
                            if (e.Pinned)
                            {
                                e.Pinned = false;
                                wasPinned = true;
                            }
                        }
                        catch
                        {
                        }

                        pinnedStates[e.Id] = wasPinned;
                    }

                    List<ConnectionSnapshot> connections = CaptureConnections(component);
                    DisconnectConnections(connections, items[0]);
                    doc.Regenerate();

                    XYZ oldAnchor = GetComponentAnchor(component);

                    foreach (Element e in component)
                    {
                        Level oldLevel = GetCurrentLevel(doc, e);
                        double oldLevelElevation = oldLevel != null ? oldLevel.Elevation : 0.0;
                        double oldOffset = GetOffsetValue(e);
                        double newOffset = oldOffset + oldLevelElevation - targetLevel.Elevation;

                        SetLevelAndOffset(e, targetLevel, newOffset);
                    }

                    doc.Regenerate();

                    XYZ newAnchor = GetComponentAnchor(component);
                    XYZ shiftBack = oldAnchor - newAnchor;

                    if (shiftBack.GetLength() > MoveTolerance)
                    {
                        List<ElementId> ids = component.Select(x => x.Id).ToList();
                        ElementTransformUtils.MoveElements(doc, ids, shiftBack);
                    }

                    doc.Regenerate();

                    RestorePointElementRotations(doc, component, poses);
                    doc.Regenerate();

                    RestoreConnections(doc, connections, items[0]);
                    doc.Regenerate();

                    bool hasValidationErrors = false;

                    foreach (TransferReportItem item in items)
                    {
                        Element committed = doc.GetElement(new ElementId(item.Id));
                        string validationError;

                        if (!ValidateElementMovedToLevel(doc, committed, targetLevel, out validationError))
                        {
                            item.AddError(validationError);
                            hasValidationErrors = true;
                        }
                    }

                    foreach (Element e in component)
                    {
                        try
                        {
                            bool restorePinned;
                            if (pinnedStates.TryGetValue(e.Id, out restorePinned) && restorePinned)
                            {
                                e.Pinned = true;
                            }
                        }
                        catch
                        {
                        }
                    }

                    if (items.Any(x => x.HasErrors) || hasValidationErrors)
                    {
                        string sharedError = string.Join(" | ",
                            items.SelectMany(x => x.GetOwnErrors())
                                 .Where(x => !string.IsNullOrWhiteSpace(x))
                                 .Distinct());

                        if (string.IsNullOrWhiteSpace(sharedError))
                            sharedError = "MEP-компонент перенесён не полностью.";

                        foreach (TransferReportItem item in items)
                        {
                            if (!item.HasErrors)
                                item.AddError("MEP-компонент перенесён не полностью. Причина: " + sharedError);
                        }

                        tx.RollBack();
                    }
                    else
                    {
                        TransactionStatus status = tx.Commit();
                        if (status != TransactionStatus.Committed)
                        {
                            string shared = "Commit вернул статус " + status + ".";
                            foreach (TransferReportItem item in items)
                                item.AddError(shared);
                        }
                        else
                        {
                            foreach (TransferReportItem item in items)
                                item.MarkSuccess("успешно.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (tx.HasStarted())
                        tx.RollBack();

                    string shared = ExceptionTextHelper.Build(ex);
                    foreach (TransferReportItem item in items)
                        item.AddError(shared);
                }
            }
        }

        private bool ValidateElementMovedToLevel(Document doc, Element element, Level targetLevel, out string errorText)
        {
            errorText = null;

            if (element == null)
            {
                errorText = "Элемент не найден после переноса.";
                return false;
            }

            if (targetLevel == null)
            {
                errorText = "Целевой уровень не найден.";
                return false;
            }

            ElementId actualLevelId;
            if (!TryGetActualAssociatedLevelId(doc, element, out actualLevelId))
            {
                errorText = "После переноса не удалось определить фактический уровень элемента.";
                return false;
            }

            if (actualLevelId == ElementId.InvalidElementId)
            {
                errorText = "После переноса фактический уровень элемента имеет InvalidElementId.";
                return false;
            }

            if (actualLevelId != targetLevel.Id)
            {
                Level actualLevel = doc.GetElement(actualLevelId) as Level;
                string actualName = actualLevel != null ? actualLevel.Name : ("ID=" + actualLevelId.IntegerValue);
                errorText = "Элемент не оказался на целевом уровне. Фактический уровень: " + actualName + ".";
                return false;
            }

            return true;
        }

        private bool ValidateGroupMovedToLevel(Document doc, Group group, Level targetLevel, out string errorText)
        {
            errorText = null;

            if (group == null)
            {
                errorText = "Группа не найдена после переноса.";
                return false;
            }

            Parameter groupLevel = GetGroupBaseLevelParameter(group);
            if (groupLevel == null || groupLevel.StorageType != StorageType.ElementId)
            {
                errorText = "После переноса не найден параметр группы \"Базовый уровень\".";
                return false;
            }

            ElementId actualLevelId = groupLevel.AsElementId();
            if (actualLevelId == ElementId.InvalidElementId)
            {
                errorText = "После переноса параметр группы \"Базовый уровень\" содержит InvalidElementId.";
                return false;
            }

            if (actualLevelId != targetLevel.Id)
            {
                Level actualLevel = doc.GetElement(actualLevelId) as Level;
                string actualName = actualLevel != null ? actualLevel.Name : ("ID=" + actualLevelId.IntegerValue);
                errorText = "Группа не оказалась на целевом уровне. Фактический уровень группы: " + actualName + ".";
                return false;
            }

            return true;
        }

        private bool TryGetActualAssociatedLevelId(Document doc, Element element, out ElementId levelId)
        {
            levelId = ElementId.InvalidElementId;

            if (element == null)
                return false;

            Group group = element as Group;
            if (group != null)
            {
                Parameter p = GetGroupBaseLevelParameter(group);
                if (p != null && p.StorageType == StorageType.ElementId)
                {
                    levelId = p.AsElementId();
                    return levelId != ElementId.InvalidElementId;
                }

                if (group.LevelId != ElementId.InvalidElementId)
                {
                    levelId = group.LevelId;
                    return true;
                }

                return false;
            }

            Parameter levelParam = GetLevelParameter(element);
            if (levelParam != null && levelParam.StorageType == StorageType.ElementId)
            {
                levelId = levelParam.AsElementId();
                if (levelId != ElementId.InvalidElementId)
                    return true;
            }

            if (element.LevelId != ElementId.InvalidElementId)
            {
                levelId = element.LevelId;
                return true;
            }

            Parameter sched = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            if (sched != null && sched.StorageType == StorageType.ElementId)
            {
                levelId = sched.AsElementId();
                if (levelId != ElementId.InvalidElementId)
                    return true;
            }

            return false;
        }

        private List<List<Element>> BuildMepConnectedComponents(List<Element> mepElements)
        {
            Dictionary<ElementId, Element> map = new Dictionary<ElementId, Element>(new ElementIdComparer());
            foreach (Element e in mepElements)
                map[e.Id] = e;

            HashSet<ElementId> visited = new HashSet<ElementId>(new ElementIdComparer());
            List<List<Element>> result = new List<List<Element>>();

            foreach (Element start in mepElements)
            {
                if (visited.Contains(start.Id))
                    continue;

                List<Element> component = new List<Element>();
                Queue<Element> queue = new Queue<Element>();
                queue.Enqueue(start);
                visited.Add(start.Id);

                while (queue.Count > 0)
                {
                    Element current = queue.Dequeue();
                    component.Add(current);

                    foreach (Connector c in GetConnectors(current))
                    {
                        if (!IsPhysicalConnector(c))
                            continue;

                        foreach (Connector refConn in GetConnectedRefs(c))
                        {
                            if (refConn == null || refConn.Owner == null)
                                continue;

                            Element owner = refConn.Owner;
                            if (owner == null)
                                continue;

                            if (!map.ContainsKey(owner.Id))
                                continue;

                            if (!visited.Contains(owner.Id))
                            {
                                visited.Add(owner.Id);
                                queue.Enqueue(owner);
                            }
                        }
                    }
                }

                result.Add(component);
            }

            return result;
        }

        private List<ConnectionSnapshot> CaptureConnections(List<Element> component)
        {
            Dictionary<string, ConnectionSnapshot> uniq = new Dictionary<string, ConnectionSnapshot>();
            HashSet<ElementId> componentIds = new HashSet<ElementId>(component.Select(x => x.Id), new ElementIdComparer());

            foreach (Element e in component)
            {
                foreach (Connector c in GetConnectors(e))
                {
                    XYZ originA;
                    if (!TryGetConnectorOrigin(c, out originA))
                        continue;

                    foreach (Connector refConn in GetConnectedRefs(c))
                    {
                        if (refConn == null || refConn.Owner == null)
                            continue;

                        XYZ originB;
                        if (!TryGetConnectorOrigin(refConn, out originB))
                            continue;

                        Element ownerB = refConn.Owner;
                        if (ownerB == null)
                            continue;

                        string key = BuildConnectionKey(e.Id, originA, ownerB.Id, originB);
                        if (uniq.ContainsKey(key))
                            continue;

                        uniq[key] = new ConnectionSnapshot
                        {
                            ElementAId = e.Id,
                            ElementBId = ownerB.Id,
                            OriginA = originA,
                            OriginB = originB,
                            DomainA = c.Domain,
                            DomainB = refConn.Domain,
                            IsExternal = !componentIds.Contains(ownerB.Id)
                        };
                    }
                }
            }

            return uniq.Values.ToList();
        }

        private void DisconnectConnections(List<ConnectionSnapshot> connections, TransferReportItem reportItem)
        {
            foreach (ConnectionSnapshot snap in connections)
            {
                try
                {
                    Connector a = FindBestConnector(snap.ElementAId, snap.OriginA, snap.DomainA);
                    Connector b = FindBestConnector(snap.ElementBId, snap.OriginB, snap.DomainB);

                    if (a != null && b != null && a.IsConnectedTo(b))
                        a.DisconnectFrom(b);
                }
                catch (Exception ex)
                {
                    if (reportItem != null)
                        reportItem.AddError("Ошибка при разъединении коннекторов: " + ExceptionTextHelper.Build(ex));
                }
            }
        }

        private void RestoreConnections(Document doc, List<ConnectionSnapshot> connections, TransferReportItem reportItem)
        {
            foreach (ConnectionSnapshot snap in connections)
            {
                try
                {
                    Connector a = FindBestConnector(snap.ElementAId, snap.OriginA, snap.DomainA);
                    Connector b = FindBestConnector(snap.ElementBId, snap.OriginB, snap.DomainB);

                    if (a == null || b == null)
                    {
                        if (reportItem != null)
                            reportItem.AddError("Не удалось найти коннекторы для восстановления соединения.");
                        continue;
                    }

                    if (!a.IsConnectedTo(b))
                        a.ConnectTo(b);
                }
                catch (Exception ex)
                {
                    if (reportItem != null)
                        reportItem.AddError("Ошибка при восстановлении соединений: " + ExceptionTextHelper.Build(ex));
                }
            }
        }

        private Connector FindBestConnector(ElementId elementId, XYZ origin, Domain domain)
        {
            Document doc = _docCache;
            if (doc == null)
                return null;

            Element e = doc.GetElement(elementId);
            if (e == null)
                return null;

            Connector best = null;
            double min = double.MaxValue;

            foreach (Connector c in GetConnectors(e))
            {
                if (c.Domain != domain)
                    continue;

                XYZ currentOrigin;
                if (!TryGetConnectorOrigin(c, out currentOrigin))
                    continue;

                double d = currentOrigin.DistanceTo(origin);
                if (d < min)
                {
                    min = d;
                    best = c;
                }
            }

            return best;
        }

        private XYZ GetComponentAnchor(List<Element> elements)
        {
            foreach (Element e in elements)
            {
                XYZ p = GetRepresentativePoint(e);
                if (p != null)
                    return p;
            }

            return XYZ.Zero;
        }

        private void RestoreElementWorldPose(Document doc, Element e, ElementPoseSnapshot beforePose)
        {
            XYZ current = GetRepresentativePoint(e);
            if (beforePose == null || current == null || beforePose.Point == null)
                return;

            XYZ delta = beforePose.Point - current;
            if (delta.GetLength() > MoveTolerance)
                ElementTransformUtils.MoveElement(doc, e.Id, delta);
        }

        private void RestorePointElementRotations(Document doc, List<Element> elements, Dictionary<ElementId, ElementPoseSnapshot> poses)
        {
            foreach (Element e in elements)
            {
                ElementPoseSnapshot snap;
                if (!poses.TryGetValue(e.Id, out snap))
                    continue;

                LocationPoint lp = e.Location as LocationPoint;
                if (lp == null || snap.HasRotation == false)
                    continue;

                double currentRotation;
                try
                {
                    currentRotation = lp.Rotation;
                }
                catch
                {
                    continue;
                }

                double delta = NormalizeAngle(snap.Rotation - currentRotation);

                if (Math.Abs(delta) < 1e-9)
                    continue;

                Line axis = Line.CreateBound(lp.Point, lp.Point + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, e.Id, axis, delta);
            }
        }

        private double NormalizeAngle(double angle)
        {
            while (angle > Math.PI) angle -= 2.0 * Math.PI;
            while (angle < -Math.PI) angle += 2.0 * Math.PI;
            return angle;
        }

        private XYZ GetRepresentativePoint(Element e)
        {
            if (e == null)
                return null;

            LocationPoint lp = e.Location as LocationPoint;
            if (lp != null)
                return lp.Point;

            LocationCurve lc = e.Location as LocationCurve;
            if (lc != null && lc.Curve != null)
                return lc.Curve.Evaluate(0.5, true);

            BoundingBoxXYZ bb = e.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return null;
        }

        private Level GetCurrentLevel(Document doc, Element e)
        {
            if (e == null)
                return null;

            Group group = e as Group;
            if (group != null)
                return GetCurrentGroupLevel(doc, group);

            Parameter p = GetLevelParameter(e);
            if (p != null && p.StorageType == StorageType.ElementId)
            {
                ElementId id = p.AsElementId();
                if (id != null && id != ElementId.InvalidElementId)
                    return doc.GetElement(id) as Level;
            }

            LevelIdData data = TryReadLevelFromKnownPlaces(doc, e);
            return data.Level;
        }

        private Level GetCurrentGroupLevel(Document doc, Group group)
        {
            if (group == null)
                return null;

            Parameter gp = GetGroupBaseLevelParameter(group);
            if (gp != null && gp.StorageType == StorageType.ElementId)
            {
                ElementId gid = gp.AsElementId();
                if (gid != null && gid != ElementId.InvalidElementId)
                {
                    Level groupLevel = doc.GetElement(gid) as Level;
                    if (groupLevel != null)
                        return groupLevel;
                }
            }

            if (group.LevelId != ElementId.InvalidElementId)
            {
                Level levelByLevelId = doc.GetElement(group.LevelId) as Level;
                if (levelByLevelId != null)
                    return levelByLevelId;
            }

            return null;
        }

        private double GetOffsetValue(Element e)
        {
            Parameter p = GetOffsetParameter(e);
            if (p != null && p.StorageType == StorageType.Double)
                return p.AsDouble();

            return 0.0;
        }

        private void SetLevelAndOffset(Element e, Level targetLevel, double newOffset)
        {
            if (e is Group)
                throw new InvalidOperationException("Для групп используйте отдельную логику переноса через Базовый уровень и Смещение начала от уровня.");

            Parameter levelParam = GetLevelParameter(e);
            if (levelParam == null)
                throw new InvalidOperationException("Не найден параметр уровня.");

            if (levelParam.IsReadOnly)
                throw new InvalidOperationException("Параметр уровня доступен только для чтения.");

            Parameter offsetParam = GetOffsetParameter(e);
            if (offsetParam == null)
                throw new InvalidOperationException("Не найден параметр смещения.");

            if (offsetParam.IsReadOnly)
                throw new InvalidOperationException("Параметр смещения доступен только для чтения.");

            levelParam.Set(targetLevel.Id);
            offsetParam.Set(newOffset);
        }

        private Parameter GetGroupBaseLevelParameter(Group group)
        {
            if (group == null)
                return null;

            Parameter p = null;

            try
            {
                p = group.LookupParameter("Базовый уровень");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.get_Parameter(BuiltInParameter.GROUP_LEVEL);
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Уровень");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Level");
                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }
            catch
            {
            }

            return null;
        }

        private Parameter GetGroupBaseOffsetParameter(Group group)
        {
            if (group == null)
                return null;

            Parameter p = null;

            try
            {
                p = group.LookupParameter("Смещение начала от уровня");
                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.get_Parameter(BuiltInParameter.GROUP_OFFSET_FROM_LEVEL);
                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Смещение от уровня");
                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Offset from Level");
                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }
            catch
            {
            }

            try
            {
                p = group.LookupParameter("Offset");
                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }
            catch
            {
            }

            return null;
        }

        private Parameter GetLevelParameter(Element e)
        {
            List<BuiltInParameter> list = new List<BuiltInParameter>
            {
                BuiltInParameter.FAMILY_LEVEL_PARAM,
                BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                BuiltInParameter.RBS_START_LEVEL_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                BuiltInParameter.LEVEL_PARAM
            };

            foreach (BuiltInParameter bip in list)
            {
                Parameter p = null;

                try
                {
                    p = e.get_Parameter(bip);
                }
                catch
                {
                    p = null;
                }

                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }

            string[] names =
            {
                "Базовый уровень",
                "Уровень",
                "Reference Level",
                "Level"
            };

            for (int i = 0; i < names.Length; i++)
            {
                Parameter p = null;

                try
                {
                    p = e.LookupParameter(names[i]);
                }
                catch
                {
                    p = null;
                }

                if (p != null && p.StorageType == StorageType.ElementId)
                    return p;
            }

            return null;
        }

        private Parameter GetOffsetParameter(Element e)
        {
            List<BuiltInParameter> list = new List<BuiltInParameter>
            {
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                BuiltInParameter.RBS_OFFSET_PARAM
            };

            foreach (BuiltInParameter bip in list)
            {
                Parameter p = null;

                try
                {
                    p = e.get_Parameter(bip);
                }
                catch
                {
                    p = null;
                }

                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }

            string[] names =
            {
                "Смещение от главной модели",
                "Смещение",
                "Смещение по высоте",
                "Смещение от уровня",
                "Смещение в середине",
                "Middle Elevation",
                "Offset",
                "Elevation from Level"
            };

            for (int i = 0; i < names.Length; i++)
            {
                Parameter p = null;

                try
                {
                    p = e.LookupParameter(names[i]);
                }
                catch
                {
                    p = null;
                }

                if (p != null && p.StorageType == StorageType.Double)
                    return p;
            }

            return null;
        }

        private LevelIdData TryReadLevelFromKnownPlaces(Document doc, Element e)
        {
            LevelIdData result = new LevelIdData();

            if (e.LevelId != ElementId.InvalidElementId)
            {
                result.Level = doc.GetElement(e.LevelId) as Level;
                if (result.Level != null)
                    return result;
            }

            Parameter sched = e.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            if (sched != null && sched.StorageType == StorageType.ElementId)
            {
                ElementId id = sched.AsElementId();
                if (id != ElementId.InvalidElementId)
                {
                    result.Level = doc.GetElement(id) as Level;
                    if (result.Level != null)
                        return result;
                }
            }

            return result;
        }

        private IEnumerable<Connector> GetConnectors(Element e)
        {
            List<Connector> result = new List<Connector>();

            MEPCurve curve = e as MEPCurve;
            if (curve != null && curve.ConnectorManager != null)
            {
                foreach (Connector c in curve.ConnectorManager.Connectors)
                    result.Add(c);

                return result;
            }

            FamilyInstance fi = e as FamilyInstance;
            if (fi != null && fi.MEPModel != null && fi.MEPModel.ConnectorManager != null)
            {
                foreach (Connector c in fi.MEPModel.ConnectorManager.Connectors)
                    result.Add(c);

                return result;
            }

            return result;
        }

        private IEnumerable<Connector> GetConnectedRefs(Connector c)
        {
            List<Connector> result = new List<Connector>();
            if (c == null || !c.IsConnected)
                return result;

            ConnectorSet refs = c.AllRefs;
            if (refs == null)
                return result;

            foreach (Connector r in refs)
            {
                if (r == null)
                    continue;

                if (r.Owner == null)
                    continue;

                if (r.Owner.Id == c.Owner.Id && r.Id == c.Id)
                    continue;

                if (!IsPhysicalConnector(r))
                    continue;

                result.Add(r);
            }

            return result;
        }

        private string BuildConnectionKey(ElementId aId, XYZ a, ElementId bId, XYZ b)
        {
            string left = aId.IntegerValue < bId.IntegerValue
                ? aId.IntegerValue + "|" + RoundPoint(a) + "|" + bId.IntegerValue + "|" + RoundPoint(b)
                : bId.IntegerValue + "|" + RoundPoint(b) + "|" + aId.IntegerValue + "|" + RoundPoint(a);

            return left;
        }

        private string RoundPoint(XYZ p)
        {
            if (p == null)
                return "0;0;0";

            return Math.Round(p.X, 6).ToString(CultureInfo.InvariantCulture) + ";"
                 + Math.Round(p.Y, 6).ToString(CultureInfo.InvariantCulture) + ";"
                 + Math.Round(p.Z, 6).ToString(CultureInfo.InvariantCulture);
        }

        private bool TryGetBuiltInCategory(Category cat, out BuiltInCategory bic)
        {
            bic = BuiltInCategory.INVALID;
            if (cat == null)
                return false;

            try
            {
                bic = (BuiltInCategory)cat.Id.IntegerValue;
                return Enum.IsDefined(typeof(BuiltInCategory), bic);
            }
            catch
            {
                return false;
            }
        }

        private bool IsPhysicalConnector(Connector c)
        {
            if (c == null)
                return false;

            try
            {
                ConnectorType type = c.ConnectorType;

                return type == ConnectorType.End
                    || type == ConnectorType.Curve
                    || type == ConnectorType.Physical;
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetConnectorOrigin(Connector c, out XYZ origin)
        {
            origin = null;

            if (c == null)
                return false;

            if (!IsPhysicalConnector(c))
                return false;

            try
            {
                origin = c.Origin;
                return origin != null;
            }
            catch
            {
                return false;
            }
        }

        private sealed class DocScope : IDisposable
        {
            private readonly ElementLevelTransferService _owner;
            private readonly Document _oldDoc;

            public DocScope(ElementLevelTransferService owner, Document doc)
            {
                _owner = owner;
                _oldDoc = owner._docCache;
                owner._docCache = doc;
            }

            public void Dispose()
            {
                _owner._docCache = _oldDoc;
            }
        }
    }

    internal sealed class LevelTransferReport
    {
        private readonly Dictionary<int, TransferReportItem> _mainItems = new Dictionary<int, TransferReportItem>();
        private readonly List<string> _globalErrors = new List<string>();

        public void AddGlobalError(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
                _globalErrors.Add(text.Trim());
        }

        public TransferReportItem GetOrCreateMainItem(Element element, Document doc)
        {
            if (element == null)
                throw new ArgumentNullException("element");

            int key = element.Id.IntegerValue;
            TransferReportItem item;
            if (_mainItems.TryGetValue(key, out item))
                return item;

            item = TransferReportItem.CreateMain(element, doc);
            _mainItems[key] = item;
            return item;
        }

        public IEnumerable<TransferReportItem> GetErrorItems()
        {
            return _mainItems.Values
                .Where(x => x != null && x.HasErrors)
                .OrderBy(x => x.Id);
        }

        public IEnumerable<TransferReportItem> GetSuccessItems()
        {
            return _mainItems.Values
                .Where(x => x != null && !x.HasErrors && x.IsSuccess)
                .OrderBy(x => x.Id);
        }

        public List<GroupedErrorEntry> BuildGroupedErrors()
        {
            Dictionary<string, GroupedErrorEntry> map = new Dictionary<string, GroupedErrorEntry>(StringComparer.Ordinal);

            foreach (string global in _globalErrors)
            {
                if (string.IsNullOrWhiteSpace(global))
                    continue;

                string key = global.Trim();
                GroupedErrorEntry entry;
                if (!map.TryGetValue(key, out entry))
                {
                    entry = new GroupedErrorEntry();
                    entry.ErrorText = key;
                    map[key] = entry;
                }

                entry.ElementLines.Add("[Глобально] " + key);
            }

            foreach (TransferReportItem item in GetErrorItems())
            {
                AppendItemErrors(map, item);
            }

            return map.Values
                .OrderByDescending(x => x.ElementLines.Count)
                .ThenBy(x => x.ErrorText, StringComparer.Ordinal)
                .ToList();
        }

        private void AppendItemErrors(Dictionary<string, GroupedErrorEntry> map, TransferReportItem item)
        {
            if (item == null)
                return;

            foreach (string err in item.GetOwnErrors())
            {
                if (string.IsNullOrWhiteSpace(err))
                    continue;

                GroupedErrorEntry entry;
                if (!map.TryGetValue(err, out entry))
                {
                    entry = new GroupedErrorEntry();
                    entry.ErrorText = err;
                    map[err] = entry;
                }

                entry.ElementLines.Add(item.BuildHeaderLine(false));
            }

            foreach (TransferReportItem child in item.GetChildren())
            {
                if (child == null)
                    continue;

                foreach (string err in child.GetOwnErrors())
                {
                    if (string.IsNullOrWhiteSpace(err))
                        continue;

                    GroupedErrorEntry entry;
                    if (!map.TryGetValue(err, out entry))
                    {
                        entry = new GroupedErrorEntry();
                        entry.ErrorText = err;
                        map[err] = entry;
                    }

                    entry.ElementLines.Add(child.BuildHeaderLine(true));
                }
            }
        }

        public List<string> BuildGroupedErrorLines()
        {
            List<string> lines = new List<string>();
            List<GroupedErrorEntry> groups = BuildGroupedErrors();

            foreach (GroupedErrorEntry group in groups)
            {
                lines.Add("Ошибка: " + group.ErrorText);

                foreach (string elementLine in group.ElementLines
                    .Distinct()
                    .OrderBy(x => x, StringComparer.Ordinal))
                {
                    lines.Add(elementLine);
                }

                lines.Add(string.Empty);
            }

            if (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1]))
                lines.RemoveAt(lines.Count - 1);

            return lines;
        }

        /// <summary>
        /// Явный список того, что не перенеслось.
        /// </summary>
        public List<string> BuildFailedLines()
        {
            List<string> lines = new List<string>();

            foreach (TransferReportItem item in GetErrorItems())
                lines.AddRange(item.BuildLines(true));

            return lines;
        }

        public List<string> BuildSuccessLines()
        {
            List<string> lines = new List<string>();

            foreach (TransferReportItem item in GetSuccessItems())
                lines.AddRange(item.BuildLines(false));

            return lines;
        }

        public string BuildText(string sourceLevelName, string targetLevelName)
        {
            List<string> lines = new List<string>();
            lines.Add("Отчёт о переносе");
            lines.Add("С уровня: " + sourceLevelName);
            lines.Add("На уровень: " + targetLevelName);

            List<string> groupedErrors = BuildGroupedErrorLines();
            if (groupedErrors.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add("Ошибки");
                lines.AddRange(groupedErrors);
            }

            List<string> failedLines = BuildFailedLines();
            if (failedLines.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add("Не перенесено");
                lines.AddRange(failedLines);
            }

            List<string> successLines = BuildSuccessLines();
            if (successLines.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add("Перенесено");
                lines.AddRange(successLines);
            }

            return string.Join(Environment.NewLine, lines);
        }
    }

    internal sealed class GroupedErrorEntry
    {
        public string ErrorText { get; set; }
        public List<string> ElementLines { get; private set; }

        public GroupedErrorEntry()
        {
            ElementLines = new List<string>();
        }
    }

    internal sealed class TransferReportItem
    {
        public int Id { get; private set; }
        public string ElementName { get; private set; }
        public string LevelParameterName { get; private set; }
        public string LevelParameterStatus { get; private set; }

        private readonly List<string> _errors = new List<string>();
        private readonly List<string> _success = new List<string>();
        private List<TransferReportItem> _children = new List<TransferReportItem>();

        public bool IsSuccess
        {
            get { return _success.Count > 0 && _errors.Count == 0; }
        }

        public bool HasErrors
        {
            get
            {
                if (_errors.Count > 0)
                    return true;

                foreach (TransferReportItem child in _children)
                {
                    if (child != null && child.HasErrors)
                        return true;
                }

                return false;
            }
        }

        public static TransferReportItem CreateMain(Element element, Document doc)
        {
            TransferReportItem item = new TransferReportItem();
            item.FillFromElement(element, doc);
            return item;
        }

        public static TransferReportItem CreateChild(Element element, Document doc)
        {
            TransferReportItem item = new TransferReportItem();
            item.FillFromElement(element, doc);
            return item;
        }

        public IList<string> GetOwnErrors()
        {
            return _errors.ToList();
        }

        public IList<TransferReportItem> GetChildren()
        {
            return _children != null
                ? _children.ToList()
                : new List<TransferReportItem>();
        }

        private void FillFromElement(Element element, Document doc)
        {
            Id = element != null ? element.Id.IntegerValue : 0;
            ElementName = ElementNamingHelper.GetElementName(element, doc);

            LevelParameterDescriptor descriptor = LevelParameterInspector.Inspect(element);
            LevelParameterName = descriptor.ParameterName;
            LevelParameterStatus = descriptor.StatusText;
        }

        public void SetChildren(IList<TransferReportItem> children)
        {
            _children = children != null
                ? children.Where(x => x != null).OrderBy(x => x.Id).ToList()
                : new List<TransferReportItem>();
        }

        public void ClearMessages()
        {
            _errors.Clear();
            _success.Clear();

            foreach (TransferReportItem child in _children)
                child.ClearMessages();
        }

        public void AddError(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            string normalized = text.Trim();

            if (!_errors.Contains(normalized))
                _errors.Add(normalized);

            _success.Clear();
        }

        public void MarkSuccess(string text)
        {
            if (HasErrors)
                return;

            if (string.IsNullOrWhiteSpace(text))
                text = "успешно.";

            string normalized = text.Trim();
            if (!_success.Contains(normalized))
                _success.Add(normalized);
        }

        public string GetCombinedErrorText()
        {
            if (_errors.Count == 0)
                return string.Empty;

            return string.Join(" | ", _errors);
        }

        public string BuildHeaderLine(bool isChild)
        {
            string prefix = isChild ? "- " : string.Empty;
            return prefix + "[" + Id + "]. " + (ElementName ?? "<Без названия>") +
                   " [" + (LevelParameterName ?? "LEVEL_PARAM") + " - " + (LevelParameterStatus ?? "НЕ НАЙДЕН") + "]";
        }

        public IEnumerable<string> BuildLines(bool forErrors)
        {
            List<string> lines = new List<string>();

            string main = BuildMainLine(forErrors);
            if (!string.IsNullOrWhiteSpace(main))
                lines.Add(main);

            foreach (TransferReportItem child in _children)
            {
                if (child == null)
                    continue;

                if (forErrors && !child.HasErrors)
                    continue;

                if (!forErrors && child.HasErrors)
                    continue;

                string childLine = child.BuildSingleLine(forErrors, true);
                if (!string.IsNullOrWhiteSpace(childLine))
                    lines.Add(childLine);
            }

            return lines;
        }

        private string BuildMainLine(bool forErrors)
        {
            return BuildSingleLine(forErrors, false);
        }

        private string BuildSingleLine(bool forErrors, bool isChild)
        {
            string head = BuildHeaderLine(isChild);

            if (forErrors)
            {
                string errorText = GetCombinedErrorText();

                if (string.IsNullOrWhiteSpace(errorText))
                    return null;

                return head + " - " + errorText;
            }
            else
            {
                if (_success.Count == 0)
                    return null;

                return head + " - " + _success[0];
            }
        }
    }

    internal sealed class LevelParameterDescriptor
    {
        public string ParameterName { get; set; }
        public string StatusText { get; set; }
    }

    internal static class LevelParameterInspector
    {
        public static LevelParameterDescriptor Inspect(Element e)
        {
            LevelParameterDescriptor descriptor = new LevelParameterDescriptor();
            descriptor.ParameterName = "LEVEL_PARAM";
            descriptor.StatusText = "НЕ НАЙДЕН";

            if (e == null)
                return descriptor;

            foreach (LevelParamCandidate candidate in GetCandidates(e))
            {
                Parameter p = null;
                try
                {
                    if (candidate.ByBuiltIn)
                        p = e.get_Parameter(candidate.BuiltInParameter);
                    else
                        p = e.LookupParameter(candidate.Name);
                }
                catch
                {
                    p = null;
                }

                if (p == null)
                    continue;

                if (p.StorageType != StorageType.ElementId)
                    continue;

                descriptor.ParameterName = candidate.DisplayName;
                descriptor.StatusText = p.IsReadOnly ? "READONLY" : "НАЙДЕН";
                return descriptor;
            }

            LevelParamCandidate fallback = GetCandidates(e).FirstOrDefault();
            if (fallback != null)
                descriptor.ParameterName = fallback.DisplayName;

            descriptor.StatusText = "НЕ НАЙДЕН";
            return descriptor;
        }

        private static List<LevelParamCandidate> GetCandidates(Element e)
        {
            List<LevelParamCandidate> list = new List<LevelParamCandidate>();

            Group g = e as Group;
            if (g != null)
            {
                list.Add(LevelParamCandidate.FromName("Базовый уровень", "Базовый уровень"));
                list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.GROUP_LEVEL, "GROUP_LEVEL"));
                list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.LEVEL_PARAM, "LEVEL_PARAM"));
                list.Add(LevelParamCandidate.FromName("Уровень", "Уровень"));
                list.Add(LevelParamCandidate.FromName("Level", "Level"));
                return list;
            }

            BuiltInCategory bic = BuiltInCategory.INVALID;
            bool hasBic = false;
            try
            {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                bic = (BuiltInCategory)e.Category.Id.IntegerValue;
#else
                bic = e.Category.BuiltInCategory;
#endif
                hasBic = true;
            }
            catch
            {
                hasBic = false;
            }

            if (hasBic && bic == BuiltInCategory.OST_PipeCurves)
            {
                list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.RBS_START_LEVEL_PARAM, "RBS_START_LEVEL_PARAM"));
                list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, "INSTANCE_REFERENCE_LEVEL_PARAM"));
                list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.LEVEL_PARAM, "LEVEL_PARAM"));
                list.Add(LevelParamCandidate.FromName("Reference Level", "Reference Level"));
                list.Add(LevelParamCandidate.FromName("Уровень", "Уровень"));
                return list;
            }

            list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.FAMILY_LEVEL_PARAM, "FAMILY_LEVEL_PARAM"));
            list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, "INSTANCE_REFERENCE_LEVEL_PARAM"));
            list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.RBS_START_LEVEL_PARAM, "RBS_START_LEVEL_PARAM"));
            list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.SCHEDULE_LEVEL_PARAM, "SCHEDULE_LEVEL_PARAM"));
            list.Add(LevelParamCandidate.FromBuiltIn(BuiltInParameter.LEVEL_PARAM, "LEVEL_PARAM"));
            list.Add(LevelParamCandidate.FromName("Базовый уровень", "Базовый уровень"));
            list.Add(LevelParamCandidate.FromName("Уровень", "Уровень"));
            list.Add(LevelParamCandidate.FromName("Reference Level", "Reference Level"));
            list.Add(LevelParamCandidate.FromName("Level", "Level"));

            return list;
        }
    }

    internal sealed class LevelParamCandidate
    {
        public bool ByBuiltIn { get; private set; }
        public BuiltInParameter BuiltInParameter { get; private set; }
        public string Name { get; private set; }
        public string DisplayName { get; private set; }

        public static LevelParamCandidate FromBuiltIn(BuiltInParameter bip, string displayName)
        {
            LevelParamCandidate c = new LevelParamCandidate();
            c.ByBuiltIn = true;
            c.BuiltInParameter = bip;
            c.DisplayName = displayName;
            return c;
        }

        public static LevelParamCandidate FromName(string name, string displayName)
        {
            LevelParamCandidate c = new LevelParamCandidate();
            c.ByBuiltIn = false;
            c.Name = name;
            c.DisplayName = displayName;
            return c;
        }
    }

    internal static class ElementNamingHelper
    {
        public static string GetElementName(Element element, Document doc)
        {
            if (element == null)
                return "<Элемент не найден>";

            Group group = element as Group;
            if (group != null)
            {
                string groupName = null;

                if (group.GroupType != null && !string.IsNullOrWhiteSpace(group.GroupType.Name))
                    groupName = group.GroupType.Name;
                else if (!string.IsNullOrWhiteSpace(group.Name))
                    groupName = group.Name;
                else
                    groupName = "<Группа>";

                return "Группа: " + groupName;
            }

            if (!string.IsNullOrWhiteSpace(element.Name))
                return element.Name;

            ElementType type = doc != null ? doc.GetElement(element.GetTypeId()) as ElementType : null;
            if (type != null && !string.IsNullOrWhiteSpace(type.Name))
                return type.Name;

            if (element.Category != null && !string.IsNullOrWhiteSpace(element.Category.Name))
                return element.Category.Name;

            return "<Без названия>";
        }
    }

    internal static class ExceptionTextHelper
    {
        public static string Build(Exception ex)
        {
            if (ex == null)
                return "Неизвестная ошибка.";

            List<string> parts = new List<string>();
            Exception current = ex;
            int guard = 0;

            while (current != null && guard < 20)
            {
                if (!string.IsNullOrWhiteSpace(current.Message))
                    parts.Add(current.Message.Trim());

                current = current.InnerException;
                guard++;
            }

            if (parts.Count == 0)
                return ex.ToString();

            return string.Join(" --> ", parts.Distinct());
        }
    }

    internal sealed class ElementPoseSnapshot
    {
        public XYZ Point { get; set; }
        public bool HasRotation { get; set; }
        public double Rotation { get; set; }

        public static ElementPoseSnapshot Create(Element e)
        {
            ElementPoseSnapshot snap = new ElementPoseSnapshot();

            if (e == null)
                return snap;

            LocationPoint lp = e.Location as LocationPoint;
            if (lp != null)
            {
                snap.Point = lp.Point;

                try
                {
                    snap.Rotation = lp.Rotation;
                    snap.HasRotation = true;
                }
                catch
                {
                    snap.HasRotation = false;
                    snap.Rotation = 0.0;
                }

                return snap;
            }

            LocationCurve lc = e.Location as LocationCurve;
            if (lc != null && lc.Curve != null)
            {
                snap.Point = lc.Curve.Evaluate(0.5, true);
                snap.HasRotation = false;
                return snap;
            }

            BoundingBoxXYZ bb = e.get_BoundingBox(null);
            if (bb != null)
            {
                snap.Point = (bb.Min + bb.Max) * 0.5;
                snap.HasRotation = false;
            }

            return snap;
        }
    }

    internal sealed class ConnectionSnapshot
    {
        public ElementId ElementAId { get; set; }
        public ElementId ElementBId { get; set; }
        public XYZ OriginA { get; set; }
        public XYZ OriginB { get; set; }
        public Domain DomainA { get; set; }
        public Domain DomainB { get; set; }
        public bool IsExternal { get; set; }
    }

    internal sealed class LevelIdData
    {
        public Level Level { get; set; }
    }

    internal class LevelElementsInfo
    {
        public List<ElementDisplayItem> Items { get; private set; }
        public int GroupCount { get; private set; }
        public int StandaloneElementCount { get; private set; }
        public int ElementCount { get; private set; }

        public LevelElementsInfo(List<ElementDisplayItem> items, int groupCount, int standaloneElementCount)
        {
            Items = items ?? new List<ElementDisplayItem>();
            GroupCount = groupCount;
            StandaloneElementCount = standaloneElementCount;
            ElementCount = Items.Count;
        }
    }

    internal class ReportWindow : Window
    {
        public ReportWindow(LevelTransferReport report, Level sourceLevel, Level targetLevel)
        {
            Title = "Отчёт переноса";
            Width = 980;
            Height = 720;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            Grid grid = new Grid();
            grid.Margin = new Thickness(10);
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            RichTextBox rtb = new RichTextBox();
            rtb.IsReadOnly = true;
            rtb.IsDocumentEnabled = false;
            rtb.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            rtb.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            rtb.Document = BuildDocument(report, sourceLevel, targetLevel);

            Grid.SetRow(rtb, 0);
            grid.Children.Add(rtb);

            Button btnCopy = new Button();
            btnCopy.Content = "Копировать в буфер";
            btnCopy.Width = 180;
            btnCopy.Height = 30;
            btnCopy.Margin = new Thickness(0, 8, 0, 0);
            btnCopy.HorizontalAlignment = HorizontalAlignment.Right;
            btnCopy.Click += delegate
            {
                try
                {
                    TextRange range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
                    Clipboard.SetText(range.Text ?? "");
                }
                catch
                {
                }
            };

            Grid.SetRow(btnCopy, 1);
            grid.Children.Add(btnCopy);

            Content = grid;
        }

        private FlowDocument BuildDocument(LevelTransferReport report, Level sourceLevel, Level targetLevel)
        {
            FlowDocument doc = new FlowDocument();
            doc.PagePadding = new Thickness(10);
            doc.FontFamily = new System.Windows.Media.FontFamily("Segoe UI");
            doc.FontSize = 13;

            AddTitle(doc, "Отчёт о переносе");

            AddLabelValue(doc, "С уровня: ", sourceLevel != null ? sourceLevel.Name : "");
            AddLabelValue(doc, "На уровень: ", targetLevel != null ? targetLevel.Name : "");

            if (report != null)
            {
                List<string> errors = report.BuildGroupedErrorLines();
                if (errors.Count > 0)
                {
                    AddSectionHeader(doc, "Ошибки");

                    foreach (string s in errors.Take(10000))
                    {
                        if (s.StartsWith("Ошибка: "))
                            AddColoredLine(doc, s, System.Windows.Media.Brushes.DarkRed);
                        else
                            AddColoredLine(doc, s, System.Windows.Media.Brushes.Red);
                    }
                }

                List<string> failed = report.BuildFailedLines();
                if (failed.Count > 0)
                {
                    AddSectionHeader(doc, "Не перенесено");

                    foreach (string s in failed.Take(10000))
                        AddColoredLine(doc, s, System.Windows.Media.Brushes.IndianRed);
                }

                List<string> success = report.BuildSuccessLines();
                if (success.Count > 0)
                {
                    AddSectionHeader(doc, "Перенесено");

                    foreach (string s in success.Take(5000))
                        AddNormalLine(doc, s);
                }
            }

            return doc;
        }

        private void AddTitle(FlowDocument doc, string text)
        {
            Paragraph p = new Paragraph();
            Run run = new Run(text);
            run.FontWeight = FontWeights.Bold;
            run.FontSize = 20;
            p.Inlines.Add(run);
            p.Margin = new Thickness(0, 0, 0, 14);
            doc.Blocks.Add(p);
        }

        private void AddSectionHeader(FlowDocument doc, string text)
        {
            Paragraph p = new Paragraph();
            Run run = new Run(text);
            run.FontWeight = FontWeights.Bold;
            run.FontSize = 16;
            p.Inlines.Add(run);
            p.Margin = new Thickness(0, 8, 0, 8);
            doc.Blocks.Add(p);
        }

        private void AddLabelValue(FlowDocument doc, string label, string value)
        {
            Paragraph p = new Paragraph();
            Run runLabel = new Run(label);
            runLabel.FontWeight = FontWeights.Bold;
            runLabel.FontSize = 14;

            Run runValue = new Run(value ?? "");
            runValue.FontSize = 14;

            p.Inlines.Add(runLabel);
            p.Inlines.Add(runValue);
            p.Margin = new Thickness(0, 0, 0, 4);
            doc.Blocks.Add(p);
        }

        private void AddColoredLine(FlowDocument doc, string text, System.Windows.Media.Brush brush)
        {
            Paragraph p = new Paragraph();
            Run run = new Run(text ?? "");
            run.Foreground = brush;
            p.Inlines.Add(run);
            p.Margin = new Thickness(0, 0, 0, 2);
            doc.Blocks.Add(p);
        }

        private void AddNormalLine(FlowDocument doc, string text)
        {
            Paragraph p = new Paragraph();
            Run run = new Run(text ?? "");
            p.Inlines.Add(run);
            p.Margin = new Thickness(0, 0, 0, 2);
            doc.Blocks.Add(p);
        }
    }

    internal class ElementDisplayItem
    {
        public ElementId ElementId { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }
        public string MainText { get; set; }
        public string GroupSuffix { get; set; }
        public int SortPriority { get; set; } // 0 - группы, 1 - элементы

        public ElementDisplayItem(Element element, Document doc)
        {
            ElementId = element.Id;
            Id = element.Id.IntegerValue;
            Name = ElementNamingHelper.GetElementName(element, doc);
            MainText = Id + " - " + Name;
            GroupSuffix = GetGroupInfo(element, doc);
            SortPriority = element is Group ? 0 : 1;
        }

        public override string ToString()
        {
            return MainText + GroupSuffix;
        }

        private static string GetGroupInfo(Element element, Document doc)
        {
            // Для самой группы не нужен хвост про группу
            if (element is Group)
                return " [Группа]";

            if (element.GroupId == ElementId.InvalidElementId)
                return string.Empty;

            Group group = doc.GetElement(element.GroupId) as Group;
            if (group == null)
                return " [Группа: <не найдена>, ID = " + element.GroupId.IntegerValue + "]";

            string groupName = string.IsNullOrWhiteSpace(group.Name) ? "(без названия)" : group.Name;
            return " [Группа: " + groupName + ", ID = " + group.Id.IntegerValue + "]";
        }
    }

    internal sealed class ElementIdComparer : IEqualityComparer<ElementId>
    {
        public bool Equals(ElementId x, ElementId y)
        {
            if (ReferenceEquals(x, y))
                return true;
            if (ReferenceEquals(x, null) || ReferenceEquals(y, null))
                return false;

            return x.IntegerValue == y.IntegerValue;
        }

        public int GetHashCode(ElementId obj)
        {
            return obj != null ? obj.IntegerValue : 0;
        }
    }

    internal sealed class TransferFailuresPreprocessor : IFailuresPreprocessor
    {
        private readonly TransferReportItem _item;

        public TransferFailuresPreprocessor(TransferReportItem item)
        {
            _item = item;
        }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
            if (failures == null || failures.Count == 0)
                return FailureProcessingResult.Continue;

            bool hasError = false;

            foreach (FailureMessageAccessor f in failures)
            {
                try
                {
                    string msg = f.GetDescriptionText();
                    if (string.IsNullOrWhiteSpace(msg))
                        msg = "Неизвестная ошибка Revit.";

                    if (_item != null)
                        _item.AddError(msg.Trim());

                    if (f.GetSeverity() == FailureSeverity.Warning)
                    {
                        try
                        {
                            failuresAccessor.DeleteWarning(f);
                        }
                        catch
                        {
                        }
                    }
                    else
                    {
                        hasError = true;
                    }
                }
                catch
                {
                }
            }

            if (hasError)
                return FailureProcessingResult.ProceedWithRollBack;

            return FailureProcessingResult.Continue;
        }
    }
}
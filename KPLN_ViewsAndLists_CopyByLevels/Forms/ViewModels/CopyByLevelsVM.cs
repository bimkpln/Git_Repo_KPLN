using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using KPLN_ViewsAndLists_Ribbon.Services.Schedules;
using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_CopyByLevels.Common;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Commands;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using KPLN_ViewsAndLists_CopyByLevels.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Input;
using FilterRule = KPLN_ViewsAndLists_Ribbon.Common.Schedules.FilterRule;

namespace KPLN_ViewsAndLists_CopyByLevels.Forms.ViewModels
{
    public sealed class CopyByLevelsVM : ObservableObject
    {
        private readonly Document _document;
        private ElementChoice _selectedSheet;
        private ElementChoice _selectedPlan;
        private CopyByLevelsM _model;
        private string _status;
        private bool _previewReady;
        private string _gap = "1.5";
        private string _startNumber = "5";
        private string _numberPrefix = "";
        private string _sheetPattern;
        private string _viewPattern = "{Исходное} — {Уровень}";

        public CopyByLevelsVM(Document document, ViewSheet preselectedSheet = null)
        {
            _document = document;
            Sheets = SourceReader.Sheets(document);
            AutoSelectCommand = new RelayCommand(AutoSelect);
            SelectAllCommand = new RelayCommand(() => SelectAll(true));
            ClearSelectionCommand = new RelayCommand(() => SelectAll(false));
            BuildPreviewCommand = new RelayCommand(() => BuildPreview());
            ViewSheet active = document.ActiveView as ViewSheet;
            SelectedSheet = Sheets.FirstOrDefault(s => preselectedSheet != null && s.Id == preselectedSheet.Id)
                ?? Sheets.FirstOrDefault(s => active != null && s.Id == active.Id)
                ?? Sheets.FirstOrDefault();
        }

        public List<ElementChoice> Sheets { get; }
        public List<ElementChoice> Plans { get; private set; } = new List<ElementChoice>();
        public CopyByLevelsM Model { get => _model; private set { _model = value; Notify(); } }
        public string Status { get => _status; private set => Change(ref _status, value); }
        public bool PreviewReady { get => _previewReady; private set => Change(ref _previewReady, value); }
        public string Gap { get => _gap; set { Change(ref _gap, value); Invalidate(); } }
        public string StartNumber { get => _startNumber; set { Change(ref _startNumber, value); Invalidate(); } }
        public string NumberPrefix { get => _numberPrefix; set { Change(ref _numberPrefix, value); Invalidate(); } }
        public string SheetPattern { get => _sheetPattern; set { Change(ref _sheetPattern, value); Invalidate(); } }
        public string ViewPattern { get => _viewPattern; set { Change(ref _viewPattern, value); Invalidate(); } }
        public ICommand AutoSelectCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand ClearSelectionCommand { get; }
        public ICommand BuildPreviewCommand { get; }

        public ElementChoice SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                if (_selectedSheet == value) return;
                _selectedSheet = value;
                Notify();
                if (value == null) return;
                Model = null;
                Plans = new List<ElementChoice>();
                _selectedPlan = null;
                Notify(nameof(Plans));
                Notify(nameof(SelectedPlan));
                try
                {
                    Model = SourceReader.Read(_document, (ViewSheet)_document.GetElement(value.Id));
                    Plans = Model.Contents.Where(c => c.View is ViewPlan && ((ViewPlan)c.View).GenLevel != null && c.View.ViewType != ViewType.AreaPlan)
                        .Select(c => new ElementChoice { Id = c.View.Id, Name = c.Name }).ToList();
                    Notify(nameof(Plans));
                    _selectedPlan = Plans.FirstOrDefault(p => p.Id == Model.MainPlan?.Id);
                    Notify(nameof(SelectedPlan));

                    int number;
                    StartNumber = int.TryParse(Model.SourceSheet.SheetNumber, out number) ? (number + 1).ToString() : "1";
                    Match prefix = Regex.Match(Model.SourceSheet.Name, @"^(.*?)План\s", RegexOptions.IgnoreCase);
                    SheetPattern = prefix.Success ? prefix.Groups[1].Value + "План {Этаж}-го этажа на отм.{Отметка}. {Типовые}"
                        : "{Исходное} — {Уровень}";
                    AutoSelect();
                    SettingsService.Load(this);
                    SubscribeRows();
                    Status = Model.MainPlan == null ? "На листе не найден этажный план. Выберите другой эталон." : "Выберите этажи и настройте состав листа. Затем на вкладке «4. Имена и предпросмотр» нажмите «Сформировать предпросмотр».";
                }
                catch (Exception ex) { Status = "Не удалось прочитать эталон: " + ex.Message; }
                Invalidate();
            }
        }

        public ElementChoice SelectedPlan
        {
            get => _selectedPlan;
            set
            {
                if (_selectedPlan == value) return;
                ElementChoice previous = _selectedPlan;
                _selectedPlan = value;
                Notify();
                Model.MainPlan = value == null ? null : _document.GetElement(value.Id) as ViewPlan;
                foreach (ContentRow content in Model.Contents)
                {
                    if (content.View.Id == value?.Id) content.Action = CopyActions.NewPlan;
                    else if (content.View.Id == previous?.Id && content.Action == CopyActions.NewPlan)
                        content.Action = content.Actions.Contains(CopyActions.Duplicate) ? CopyActions.Duplicate : CopyActions.Skip;
                }
                SourceReader.PopulateLevels(Model);
                AutoSelect();
                SubscribeRows();
                Invalidate();
            }
        }

        private void SubscribeRows()
        {
            Model.PlanDetails.PropertyChanged -= PlanDetailsChanged;
            Model.PlanDetails.PropertyChanged += PlanDetailsChanged;
            foreach (LevelRow row in Model.Levels) row.PropertyChanged += (s, e) => Invalidate();
            foreach (ContentRow row in Model.Contents) row.PropertyChanged += (s, e) => Invalidate();
            foreach (ScheduleRow row in Model.Schedules)
            {
                row.PropertyChanged += (s, e) => Invalidate();
                foreach (FilterRule filter in row.Filters) filter.PropertyChanged += (s, e) => Invalidate();
            }
        }

        private void PlanDetailsChanged(object sender, System.ComponentModel.PropertyChangedEventArgs args)
        {
            Invalidate();
        }

        public void Invalidate()
        {
            if (PreviewReady)
                Status = "Настройки изменены. На вкладке «4. Имена и предпросмотр» нажмите «Сформировать предпросмотр» повторно.";
            PreviewReady = false;
        }

        private void SelectAll(bool selected)
        {
            if (Model == null) return;
            foreach (LevelRow row in Model.Levels) row.Selected = selected;
            Invalidate();
        }

        private void AutoSelect()
        {
            if (Model?.MainPlan?.GenLevel == null) return;
            double gap;
            if (!double.TryParse((Gap ?? "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out gap) || gap <= 0)
            {
                Status = "Введите положительный шаг уровней в метрах.";
                return;
            }
            bool[] selection = NamingRules.SelectLevels(Model.Levels.Select(l => l.Height).ToList(), Model.MainPlan.GenLevel.ProjectElevation * 0.3048, gap);
            for (int i = 0; i < selection.Length; i++) Model.Levels[i].Selected = selection[i];
            Invalidate();
        }

        public bool BuildPreview()
        {
            PreviewReady = false;
            if (Model == null) return false;
            Model.Preview.Clear();
            try
            {
                if (Model.MainPlan == null) throw new InvalidOperationException("Выберите основной план эталона.");
                int start;
                if (!int.TryParse(StartNumber, out start) || start < 0) throw new InvalidOperationException("Начальный номер листа должен быть целым неотрицательным числом.");
                var levels = Model.Levels.Where(l => l.Selected).ToList();
                if (levels.Count == 0) throw new InvalidOperationException("Не выбран ни один уровень.");

                ContentRow mainContent = Model.Contents.FirstOrDefault(c => c.View.Id == Model.MainPlan.Id);
                bool useExistingMainView = mainContent?.Action == CopyActions.Replace;
                View existingMainView = useExistingMainView && mainContent.Replacement != null
                    ? _document.GetElement(mainContent.Replacement.Id) as View : null;

                int index = 0;
                foreach (LevelRow level in levels)
                {
                    index++;
                    Model.Preview.Add(new PreviewRow
                    {
                        Level = level, Index = index, Number = NumberPrefix + (start + index - 1).ToString(CultureInfo.InvariantCulture),
                        Name = NamingRules.Expand(SheetPattern, level.Floor, level.Name, level.Elevation, level.Group, Model.SourceSheet.Name, index),
                        ViewName = useExistingMainView ? existingMainView?.Name ?? ""
                            : NamingRules.Expand(ViewPattern, level.Floor, level.Name, level.Elevation, level.Group, Model.MainPlan.Name, index),
                        Filters = string.Join("; ", Model.Schedules.Where(s => s.Include && s.Copy)
                            .SelectMany(s => ScheduleLevelFilterService.Describe(s, level, index)))
                    });
                }
                ValidatePreview();
                return PreviewReady;
            }
            catch (Exception ex)
            {
                foreach (PreviewRow row in Model.Preview.Where(row => row.Status == "Ожидает проверки"))
                    row.Status = "Проверка не пройдена";
                Status = "Создание недоступно: " + ex.Message;
                return false;
            }
        }

        public bool ValidatePreview()
        {
            PreviewReady = false;
            try
            {
                if (Model.Preview.Count == 0) throw new InvalidOperationException("Сначала сформируйте предпросмотр.");
                ContentRow mainContent = Model.Contents.FirstOrDefault(c => c.View.Id == Model.MainPlan.Id);
                if (mainContent == null || (mainContent.Action != CopyActions.NewPlan && mainContent.Action != CopyActions.Replace))
                    throw new InvalidOperationException("Для основного плана выберите «Новый план уровня» или «Другой существующий вид».");

                bool createMainPlan = mainContent.Action == CopyActions.NewPlan;
                View existingMainView = createMainPlan || mainContent.Replacement == null
                    ? null : _document.GetElement(mainContent.Replacement.Id) as View;
                if (!createMainPlan && existingMainView == null)
                    throw new InvalidOperationException("Выберите существующий вид для замены основного плана.");

                var sheets = new FilteredElementCollector(_document).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().ToList();
                var viewNames = new HashSet<string>(new FilteredElementCollector(_document).OfClass(typeof(View)).Cast<View>().Select(v => v.Name), StringComparer.OrdinalIgnoreCase);
                var numbers = new HashSet<string>(sheets.Select(s => s.SheetNumber), StringComparer.OrdinalIgnoreCase);
                var errors = new List<string>();
                foreach (PreviewRow row in Model.Preview)
                {
                    row.SkipExisting = Model.SkipExisting && sheets.Any(s => GenerationService.WasCreated(s, Model.SourceSheet, row.Level.Level));
                    row.Status = row.SkipExisting ? "Пропуск: уже создан" : "Ожидает проверки";
                    if (row.SkipExisting) continue;
                    // Существующий вид размещается без переименования и проверки имени как нового.
                    if (!createMainPlan) row.ViewName = existingMainView.Name;
                    string error = RevitHelpers.InvalidName(row.Name) ?? RevitHelpers.InvalidName(row.Number);
                    if (createMainPlan) error = error ?? RevitHelpers.InvalidName(row.ViewName);
                    if (string.IsNullOrWhiteSpace(row.Level.Floor)) error = "Не задано обозначение этажа";
                    if (!numbers.Add(row.Number)) error = "Номер листа уже занят: " + row.Number;
                    if (createMainPlan && !viewNames.Add(row.ViewName)) error = "Имя плана уже занято: " + row.ViewName;
                    if (error != null) { row.Status = error; errors.Add(row.LevelName + ": " + error); }
                }

                int count = Model.Preview.Count(r => !r.SkipExisting);
                foreach (ContentRow content in Model.Contents.Where(c => c.Action != CopyActions.Skip))
                {
                    if (!content.Actions.Contains(content.Action)) throw new InvalidOperationException("Недоступное действие для вида: " + content.Name);
                    if (content.Action == CopyActions.Replace)
                    {
                        View replacement = content.Replacement == null ? null : _document.GetElement(content.Replacement.Id) as View;
                        if (replacement == null) throw new InvalidOperationException("Выберите замену для вида: " + content.Name);
                        if (count > 1 && replacement.ViewType != ViewType.Legend)
                            throw new InvalidOperationException("Один обычный вид нельзя разместить на нескольких листах: " + replacement.Name + ". Выберите независимые копии или создайте один лист.");
                        if (replacement.ViewType != ViewType.Legend && new FilteredElementCollector(_document).OfClass(typeof(Viewport))
                            .Cast<Viewport>().Any(p => p.ViewId == replacement.Id))
                            throw new InvalidOperationException("Вид для замены уже размещён на листе: " + replacement.Name);
                    }
                    if (content.Action == CopyActions.NewPlan)
                    {
                        var plan = (ViewPlan)content.View;
                        using (var crop = plan.GetCropRegionShapeManager())
                            if (crop.NumberOfSplitRegions > 1) throw new InvalidOperationException("Разделённая обрезка не поддерживается: " + plan.Name);
                    }
                }

                foreach (ScheduleRow schedule in Model.Schedules.Where(s => s.Include))
                {

                    if (!schedule.Copy && ScheduleLevelFilterService.HasChanges(schedule))
                        throw new InvalidOperationException("Для изменения фильтров включите независимую копию: " + schedule.Name);
                    if (schedule.LevelFilters != null)
                    {
                        ScheduleLevelFilterService.Validate(_document, schedule, schedule.LevelFilters,
                            Model.Preview.Where(row => !row.SkipExisting).Select(row => new ScheduleOperation
                            {
                                Level = row.Level, Index = row.Index
                            }));
                    }
                    else
                    {
                        foreach (PreviewRow row in Model.Preview.Where(r => !r.SkipExisting))
                            foreach (FilterRule rule in schedule.Filters)
                                ScheduleService.ResolveFilter(_document, schedule.Schedule, rule, row.Level, row.Index);
                    }
                }

                if (errors.Count > 0) throw new InvalidOperationException(string.Join("\r\n", errors));
                foreach (PreviewRow row in Model.Preview.Where(row => !row.SkipExisting))
                    row.Status = "Создать";
                PreviewReady = true;
                int unchanged = Model.Schedules.Count(s => s.Include && s.Copy && !ScheduleLevelFilterService.HasChanges(s));
                Status = "К созданию: " + count + "; пропусков: " + (Model.Preview.Count - count) + "."
                    + (unchanged > 0 ? " ВНИМАНИЕ: у " + unchanged + " копий спецификаций все фильтры сохранены — проверьте, что они покажут нужный этаж." : "")
                    + (Model.Contents.Any(c => c.Action == CopyActions.NewPlan)
                        ? " Марки и размеры нового этажа требуют отдельного оформления."
                        : " Выбранные существующие виды будут размещены без изменения их имён и оформления.");
                return true;
            }
            catch (Exception ex)
            {
                foreach (PreviewRow row in Model.Preview.Where(row => row.Status == "Ожидает проверки"))
                    row.Status = "Проверка не пройдена";
                Status = "Создание недоступно: " + ex.Message;
                return false;
            }
        }
    }
}

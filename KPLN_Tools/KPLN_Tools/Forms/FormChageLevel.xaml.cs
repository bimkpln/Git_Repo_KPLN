using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Visibility = System.Windows.Visibility;

namespace KPLN_Tools.Forms
{
    public partial class FormChageLevel : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;

        private readonly List<BuiltInCategory> _listCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_AreaRein,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Rebar,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_GenericModel
        };

        private readonly List<BuiltInCategory> _editableCategories = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_StructuralColumns
        };

        private List<Level> _levels;
        private List<Element> _currentLevelElements;

        public ObservableCollection<CategoryOffsetItem> CategoryItems { get; private set; }

        private ObservableCollection<CategoryOffsetItem> _visibleCategoryItems;
        public ObservableCollection<CategoryOffsetItem> VisibleCategoryItems
        {
            get { return _visibleCategoryItems; }
            set
            {
                _visibleCategoryItems = value;
                OnPropertyChanged("VisibleCategoryItems");
            }
        }

        public FormChageLevel(Document document)
        {
            _doc = document;

            InitializeComponent();
            DataContext = this;

            BuildCategoryItems();
            VisibleCategoryItems = new ObservableCollection<CategoryOffsetItem>();
            LoadData();
        }

        public class CategoryOffsetItem : INotifyPropertyChanged
        {
            private int _count;
            private string _firstOffsetText;

            public BuiltInCategory Category { get; set; }
            public string DisplayName { get; set; }

            public bool HasFirstOffset { get; set; }
            public string FirstOffsetLabel { get; set; }

            public Visibility FirstOffsetVisibility
            {
                get { return HasFirstOffset ? Visibility.Visible : Visibility.Collapsed; }
            }

            public int Count
            {
                get { return _count; }
                set
                {
                    _count = value;
                    OnPropertyChanged("Count");
                }
            }

            public string FirstOffsetText
            {
                get { return _firstOffsetText; }
                set
                {
                    _firstOffsetText = value;
                    OnPropertyChanged("FirstOffsetText");
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged(string propertyName)
            {
                PropertyChangedEventHandler handler = PropertyChanged;
                if (handler != null)
                {
                    handler(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }

        private class OffsetValues
        {
            public bool HasFirstOffset;
            public double FirstOffsetDelta;
        }

        private class FailedElementInfo
        {
            public string CategoryName { get; set; }
            public int ElementId { get; set; }
            public string Reason { get; set; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private void BuildCategoryItems()
        {
            CategoryItems = new ObservableCollection<CategoryOffsetItem>
            {
                new CategoryOffsetItem
                {
                    Category = BuiltInCategory.OST_StructuralFraming,
                    DisplayName = "Каркас несущий",
                    HasFirstOffset = true,
                    FirstOffsetLabel = "Смещение:",
                    FirstOffsetText = "0"
                },
                new CategoryOffsetItem
                {
                    Category = BuiltInCategory.OST_Floors,
                    DisplayName = "Перекрытия",
                    HasFirstOffset = true,
                    FirstOffsetLabel = "Смещение от уровня:",
                    FirstOffsetText = "0"
                },
                new CategoryOffsetItem
                {
                    Category = BuiltInCategory.OST_Walls,
                    DisplayName = "Стены",
                    HasFirstOffset = true,
                    FirstOffsetLabel = "Смещение:",
                    FirstOffsetText = "0"
                },
                new CategoryOffsetItem
                {
                    Category = BuiltInCategory.OST_Windows,
                    DisplayName = "Окна",
                    HasFirstOffset = true,
                    FirstOffsetLabel = "Высота нижнего бруса:",
                    FirstOffsetText = "0"
                },
                new CategoryOffsetItem
                {
                    Category = BuiltInCategory.OST_StructuralColumns,
                    DisplayName = "Несущие колонны",
                    HasFirstOffset = true,
                    FirstOffsetLabel = "Смещение:",
                    FirstOffsetText = "0"
                }
            };
        }

        private void LoadData()
        {
            _levels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .ToList();

            LevelSelector.ItemsSource = _levels;

            if (_levels.Count > 0)
            {
                LevelSelector.SelectedIndex = 0;
            }
            else
            {
                RefreshForSelectedLevel();
            }
        }

        private static BuiltInCategory GetBuiltInCategory(Category category)
        {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
            return (BuiltInCategory)category.Id.IntegerValue;
#else
            return category.BuiltInCategory;
#endif
        }

        private ElementId GetElementBaseLevelId(Element element)
        {
            if (element == null || element.Category == null)
            {
                return ElementId.InvalidElementId;
            }

            BuiltInCategory bic = GetBuiltInCategory(element.Category);

            switch (bic)
            {
                case BuiltInCategory.OST_Walls:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.WALL_BASE_CONSTRAINT);

                case BuiltInCategory.OST_StructuralColumns:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                        BuiltInParameter.SCHEDULE_BASE_LEVEL_PARAM,
                        BuiltInParameter.LEVEL_PARAM);

                case BuiltInCategory.OST_Floors:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.LEVEL_PARAM);

                case BuiltInCategory.OST_Windows:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.FAMILY_LEVEL_PARAM,
                        BuiltInParameter.LEVEL_PARAM);

                case BuiltInCategory.OST_StructuralFraming:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                        BuiltInParameter.LEVEL_PARAM);

                case BuiltInCategory.OST_GenericModel:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.FAMILY_LEVEL_PARAM,
                        BuiltInParameter.LEVEL_PARAM);

                case BuiltInCategory.OST_Rebar:
                case BuiltInCategory.OST_AreaRein:
                    return GetElementIdParameterValue(
                        element,
                        BuiltInParameter.LEVEL_PARAM);

                default:
                    return ElementId.InvalidElementId;
            }
        }

        private ElementId GetElementIdParameterValue(Element element, params BuiltInParameter[] parameterIds)
        {
            foreach (BuiltInParameter parameterId in parameterIds)
            {
                Parameter parameter = element.get_Parameter(parameterId);
                if (parameter == null || !parameter.HasValue)
                {
                    continue;
                }

                try
                {
                    ElementId value = parameter.AsElementId();
                    if (value != ElementId.InvalidElementId)
                    {
                        return value;
                    }
                }
                catch
                {
                }
            }

            return ElementId.InvalidElementId;
        }

        private Parameter GetFirstWritableParameter(Element element, params BuiltInParameter[] parameterIds)
        {
            foreach (BuiltInParameter parameterId in parameterIds)
            {
                Parameter parameter = element.get_Parameter(parameterId);
                if (parameter != null && !parameter.IsReadOnly)
                {
                    return parameter;
                }
            }

            return null;
        }

        private bool HasValidTopConstraint(Element wall)
        {
            Parameter topConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            if (topConstraintParam == null || !topConstraintParam.HasValue)
            {
                return false;
            }

            try
            {
                ElementId topConstraintId = topConstraintParam.AsElementId();
                return topConstraintId != ElementId.InvalidElementId;
            }
            catch
            {
                return false;
            }
        }

        private bool IsWallRelatedToLevelForInfo(Element wall, ElementId levelId)
        {
            if (wall == null)
            {
                return false;
            }

            ElementId baseLevelId = GetElementIdParameterValue(
                wall,
                BuiltInParameter.WALL_BASE_CONSTRAINT);

            if (baseLevelId == levelId)
            {
                return true;
            }

            ElementId topLevelId = GetElementIdParameterValue(
                wall,
                BuiltInParameter.WALL_HEIGHT_TYPE);

            if (topLevelId == levelId)
            {
                return true;
            }

            return false;
        }

        private void LevelSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshForSelectedLevel();
        }

        private void RefreshForSelectedLevel()
        {
            RefreshElementList();
            RefreshCategoryCounts();
            RefreshVisibleCategoryItems();
        }

        private void RefreshElementList()
        {
            ElementList.Document.Blocks.Clear();

            Level selectedLevel = LevelSelector.SelectedItem as Level;
            if (selectedLevel == null)
            {
                ElementList.Document.Blocks.Add(new Paragraph(new Run("Уровень не выбран.")));
                _currentLevelElements = new List<Element>();
                return;
            }

            _currentLevelElements = CollectElementsByLevel(selectedLevel.Id);

            if (_currentLevelElements.Count == 0)
            {
                ElementList.Document.Blocks.Add(new Paragraph(new Run("На выбранном уровне нет элементов поддерживаемых категорий.")));
                return;
            }

            StringBuilder sb = new StringBuilder();

            foreach (Element element in _currentLevelElements.OrderBy(x => x.Id.IntegerValue))
            {
                string categoryName = element.Category != null ? element.Category.Name : "<без категории>";
                sb.AppendLine("ID: " + element.Id.IntegerValue + "; Категория: " + categoryName + "; Имя: " + element.Name);
            }

            ElementList.Document.Blocks.Add(new Paragraph(new Run(sb.ToString())));
        }

        private void RefreshCategoryCounts()
        {
            foreach (CategoryOffsetItem item in CategoryItems)
            {
                item.Count = 0;
            }

            if (_currentLevelElements == null)
            {
                return;
            }

            Level selectedLevel = LevelSelector.SelectedItem as Level;
            if (selectedLevel == null)
            {
                return;
            }

            foreach (Element element in _currentLevelElements)
            {
                if (element == null || element.Category == null)
                {
                    continue;
                }

                BuiltInCategory bic = GetBuiltInCategory(element.Category);

                if (bic == BuiltInCategory.OST_Walls)
                {
                    continue;
                }

                CategoryOffsetItem item = CategoryItems.FirstOrDefault(x => x.Category == bic);
                if (item != null)
                {
                    item.Count++;
                }
            }

            CategoryOffsetItem wallItem = CategoryItems.FirstOrDefault(x => x.Category == BuiltInCategory.OST_Walls);
            if (wallItem != null)
            {
                List<Element> allWalls = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                wallItem.Count = allWalls.Count(x => IsWallRelatedToLevelForInfo(x, selectedLevel.Id));
            }
        }

        private void RefreshVisibleCategoryItems()
        {
            List<CategoryOffsetItem> sortedVisibleItems = CategoryItems
                .Where(x => x.Count > 0)
                .OrderByDescending(x => x.Count)
                .ThenBy(x => x.DisplayName)
                .ToList();

            VisibleCategoryItems = new ObservableCollection<CategoryOffsetItem>(sortedVisibleItems);
        }

        private List<Element> CollectElementsByLevel(ElementId levelId)
        {
            List<Element> result = new List<Element>();

            List<ElementFilter> categoryFilters = _listCategories
                .Select(x => (ElementFilter)new ElementCategoryFilter(x))
                .ToList();

            if (categoryFilters.Count == 0)
            {
                return result;
            }

            LogicalOrFilter multiCategoryFilter = new LogicalOrFilter(categoryFilters);

            List<Element> elements = new FilteredElementCollector(_doc)
                .WherePasses(multiCategoryFilter)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            foreach (Element element in elements)
            {
                ElementId baseLevelId = GetElementBaseLevelId(element);
                if (baseLevelId == levelId)
                {
                    result.Add(element);
                }
            }

            return result;
        }

        private void Button_ClickCopyValues(object sender, RoutedEventArgs e)
        {
            CopyValuesToAllFields();
        }

        private void CopyValuesToAllFields()
        {
            CategoryOffsetItem sourceItem = CategoryItems.FirstOrDefault(x => IsNonZeroText(x.FirstOffsetText));

            if (sourceItem == null)
            {
                TaskDialog.Show("Предупреждение", "Не найдено ни одного поля со значением, отличным от нуля.");
                return;
            }

            string valueToCopy = sourceItem.FirstOffsetText;

            foreach (CategoryOffsetItem item in CategoryItems)
            {
                if (item.HasFirstOffset)
                {
                    item.FirstOffsetText = valueToCopy;
                }
            }
        }

        private bool IsNonZeroText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            double value;
            if (!TryParseLength(text, out value))
            {
                return false;
            }

            return Math.Abs(value) > 1e-9;
        }

        private void Button_ClickApplyOffsets(object sender, RoutedEventArgs e)
        {
            ApplyOffsets();
        }

        private void ApplyOffsets()
        {
            Level selectedLevel = LevelSelector.SelectedItem as Level;
            if (selectedLevel == null)
            {
                TaskDialog.Show("Предупреждение", "Не выбран уровень.");
                return;
            }

            if (_currentLevelElements == null || _currentLevelElements.Count == 0)
            {
                TaskDialog.Show("Предупреждение", "На выбранном уровне нет элементов поддерживаемых категорий.");
                return;
            }

            Dictionary<BuiltInCategory, OffsetValues> offsetsByCategory;
            string validationMessage;

            if (!TryBuildOffsetsDictionary(out offsetsByCategory, out validationMessage))
            {
                TaskDialog.Show("Ошибка ввода", validationMessage);
                return;
            }

            int movedCount = 0;
            List<FailedElementInfo> failedItems = new List<FailedElementInfo>();

            using (Transaction transaction = new Transaction(_doc, "KPLN: Смещение элементов выбранного уровня"))
            {
                transaction.Start();

                foreach (Element element in _currentLevelElements)
                {
                    if (element == null || element.Category == null)
                    {
                        failedItems.Add(new FailedElementInfo
                        {
                            CategoryName = "Неизвестная категория",
                            ElementId = element != null ? element.Id.IntegerValue : -1,
                            Reason = "Элемент не найден или отсутствует категория."
                        });
                        continue;
                    }

                    BuiltInCategory bic = GetBuiltInCategory(element.Category);

                    if (!_editableCategories.Contains(bic))
                    {
                        continue;
                    }

                    OffsetValues offsets;
                    if (!offsetsByCategory.TryGetValue(bic, out offsets))
                    {
                        continue;
                    }

                    bool hasDelta = offsets.HasFirstOffset && Math.Abs(offsets.FirstOffsetDelta) > 1e-9;
                    if (!hasDelta)
                    {
                        continue;
                    }

                    string failReason;
                    try
                    {
                        bool moved = ApplyOffsetsToElement(element, offsets, out failReason);

                        if (moved)
                        {
                            movedCount++;
                        }
                        else
                        {
                            failedItems.Add(new FailedElementInfo
                            {
                                CategoryName = element.Category.Name,
                                ElementId = element.Id.IntegerValue,
                                Reason = failReason
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        failedItems.Add(new FailedElementInfo
                        {
                            CategoryName = element.Category.Name,
                            ElementId = element.Id.IntegerValue,
                            Reason = "Исключение: " + ex.Message
                        });
                    }
                }

                transaction.Commit();
            }

            RefreshForSelectedLevel();
            ShowResultDialog(movedCount, failedItems);

            Activate();
            Show();
        }

        private bool TryBuildOffsetsDictionary(out Dictionary<BuiltInCategory, OffsetValues> offsetsByCategory, out string validationMessage)
        {
            offsetsByCategory = new Dictionary<BuiltInCategory, OffsetValues>();
            validationMessage = null;

            foreach (CategoryOffsetItem item in CategoryItems)
            {
                OffsetValues values = new OffsetValues
                {
                    HasFirstOffset = item.HasFirstOffset,
                    FirstOffsetDelta = 0.0
                };

                if (item.HasFirstOffset)
                {
                    double parsedOffset;
                    if (!TryParseLength(item.FirstOffsetText, out parsedOffset))
                    {
                        validationMessage = "Не удалось распознать значение поля:\n" +
                                            item.FirstOffsetLabel +
                                            "\nдля категории:\n" +
                                            item.DisplayName;
                        return false;
                    }

                    values.FirstOffsetDelta = parsedOffset;
                }

                offsetsByCategory[item.Category] = values;
            }

            return true;
        }

        private bool ApplyOffsetsToElement(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;

            if (element == null || element.Category == null)
            {
                failReason = "Элемент отсутствует или не имеет категории.";
                return false;
            }

            BuiltInCategory bic = GetBuiltInCategory(element.Category);

            switch (bic)
            {
                case BuiltInCategory.OST_StructuralFraming:
                    return ApplyOffsetsToStructuralFraming(element, offsets, out failReason);

                case BuiltInCategory.OST_Floors:
                    return ApplyOffsetsToFloor(element, offsets, out failReason);

                case BuiltInCategory.OST_Walls:
                    return ApplyOffsetsToWall(element, offsets, out failReason);

                case BuiltInCategory.OST_Windows:
                    return ApplyOffsetsToWindow(element, offsets, out failReason);

                case BuiltInCategory.OST_StructuralColumns:
                    return ApplyOffsetsToColumn(element, offsets, out failReason);

                default:
                    failReason = "Категория не поддерживается для смещения.";
                    return false;
            }
        }

        private bool ApplyOffsetsToStructuralFraming(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;
            double delta = offsets.FirstOffsetDelta;

            if (Math.Abs(delta) < 1e-9)
            {
                return true;
            }

            Parameter startOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);

            if (startOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи поле STRUCTURAL_BEAM_END0_ELEVATION.";
                return false;
            }

            Parameter endOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);

            if (endOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи поле STRUCTURAL_BEAM_END1_ELEVATION.";
                return false;
            }

            startOffsetParam.Set(startOffsetParam.AsDouble() + delta);
            endOffsetParam.Set(endOffsetParam.AsDouble() + delta);
            return true;
        }

        private bool ApplyOffsetsToFloor(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;
            double delta = offsets.FirstOffsetDelta;

            if (Math.Abs(delta) < 1e-9)
            {
                return true;
            }

            Parameter floorOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);

            if (floorOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи поле FLOOR_HEIGHTABOVELEVEL_PARAM.";
                return false;
            }

            floorOffsetParam.Set(floorOffsetParam.AsDouble() + delta);
            return true;
        }

        private bool ApplyOffsetsToWall(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;
            double delta = offsets.FirstOffsetDelta;

            if (Math.Abs(delta) < 1e-9)
            {
                return true;
            }

            Parameter baseOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.WALL_BASE_OFFSET);

            if (baseOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи поле WALL_BASE_OFFSET.";
                return false;
            }

            baseOffsetParam.Set(baseOffsetParam.AsDouble() + delta);

            bool hasTopConstraint = HasValidTopConstraint(element);

            if (hasTopConstraint)
            {
                Parameter topOffsetParam = GetFirstWritableParameter(
                    element,
                    BuiltInParameter.WALL_TOP_OFFSET);

                if (topOffsetParam == null)
                {
                    failReason = "Не найдено доступное для записи поле WALL_TOP_OFFSET.";
                    return false;
                }

                topOffsetParam.Set(topOffsetParam.AsDouble() + delta);
                return true;
            }

            Parameter unconnectedHeightParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.WALL_USER_HEIGHT_PARAM);

            if (unconnectedHeightParam == null)
            {
                failReason = "Не найдено доступное для записи поле WALL_USER_HEIGHT_PARAM для неприсоединённой стены.";
                return false;
            }

            double newHeight = unconnectedHeightParam.AsDouble() + delta;
            if (newHeight < 1e-6)
            {
                failReason = "Неприсоединённая высота стала меньше допустимой.";
                return false;
            }

            unconnectedHeightParam.Set(newHeight);
            return true;
        }

        private bool ApplyOffsetsToWindow(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;
            double delta = offsets.FirstOffsetDelta;

            if (Math.Abs(delta) < 1e-9)
            {
                return true;
            }

            Parameter sillHeightParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);

            if (sillHeightParam == null)
            {
                failReason = "Не найдено доступное для записи поле INSTANCE_SILL_HEIGHT_PARAM.";
                return false;
            }

            sillHeightParam.Set(sillHeightParam.AsDouble() + delta);
            return true;
        }

        private bool ApplyOffsetsToColumn(Element element, OffsetValues offsets, out string failReason)
        {
            failReason = null;
            double delta = offsets.FirstOffsetDelta;

            if (Math.Abs(delta) < 1e-9)
            {
                return true;
            }

            Parameter baseOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM);

            if (baseOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи нижнее смещение колонны.";
                return false;
            }

            Parameter topOffsetParam = GetFirstWritableParameter(
                element,
                BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
                BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM);

            if (topOffsetParam == null)
            {
                failReason = "Не найдено доступное для записи верхнее смещение колонны.";
                return false;
            }

            baseOffsetParam.Set(baseOffsetParam.AsDouble() + delta);
            topOffsetParam.Set(topOffsetParam.AsDouble() + delta);
            return true;
        }


        private bool TryParseLength(string input, out double internalValue)
        {
            internalValue = 0.0;

            if (string.IsNullOrWhiteSpace(input))
            {
                return true;
            }

            string normalized = input.Trim().Replace(',', '.');

            double numericValue;
            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out numericValue))
            {
#if Revit2020 || Debug2020
                return UnitFormatUtils.TryParse(_doc.GetUnits(), UnitType.UT_Length, input, out internalValue);
#else
                return UnitFormatUtils.TryParse(_doc.GetUnits(), SpecTypeId.Length, input, out internalValue);
#endif
            }

#if Revit2020 || Debug2020
            FormatOptions formatOptions = _doc.GetUnits().GetFormatOptions(UnitType.UT_Length);
            DisplayUnitType displayUnitType = formatOptions.DisplayUnits;
            internalValue = UnitUtils.ConvertToInternalUnits(numericValue, displayUnitType);
#else
            FormatOptions formatOptions = _doc.GetUnits().GetFormatOptions(SpecTypeId.Length);
            ForgeTypeId unitTypeId = formatOptions.GetUnitTypeId();
            internalValue = UnitUtils.ConvertToInternalUnits(numericValue, unitTypeId);
#endif
            return true;
        }

        private void ShowResultDialog(int movedCount, List<FailedElementInfo> failedItems)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Успешно обработано элементов: " + movedCount);

            if (failedItems.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Есть ошибки: " + failedItems.Count);
                sb.AppendLine("Сохранить лог?");
            }

            if (failedItems.Count == 0)
            {
                TaskDialog.Show("Результат смещения", sb.ToString());
                return;
            }

            TaskDialog dialog = new TaskDialog("Результат смещения");
            dialog.MainInstruction = "Смещение выполнено с ошибками.";
            dialog.MainContent = sb.ToString();
            dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            dialog.DefaultButton = TaskDialogResult.Yes;

            TaskDialogResult result = dialog.Show();

            if (result == TaskDialogResult.Yes)
            {
                SaveFailureLog(failedItems);
            }
        }

        private void SaveFailureLog(List<FailedElementInfo> failedItems)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Title = "Сохранить лог смещения";
            saveDialog.Filter = "Текстовый файл (*.txt)|*.txt";
            saveDialog.FileName = "KPLN_СмещениеЭлементов_Лог.txt";

            bool? dialogResult = saveDialog.ShowDialog();
            if (dialogResult != true)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Лог смещения элементов");
            sb.AppendLine("Документ: " + _doc.Title);
            sb.AppendLine("Дата: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine();

            foreach (IGrouping<string, FailedElementInfo> categoryGroup in failedItems
                .OrderBy(x => x.CategoryName)
                .ThenBy(x => x.Reason)
                .ThenBy(x => x.ElementId)
                .GroupBy(x => x.CategoryName))
            {
                sb.AppendLine("Категория: " + categoryGroup.Key);

                foreach (IGrouping<string, FailedElementInfo> reasonGroup in categoryGroup.GroupBy(x => x.Reason))
                {
                    List<int> ids = reasonGroup
                        .Select(x => x.ElementId)
                        .Distinct()
                        .OrderBy(x => x)
                        .ToList();

                    sb.AppendLine("  Причина: " + reasonGroup.Key);
                    sb.AppendLine("  ID: " + string.Join(", ", ids));
                    sb.AppendLine();
                }

                sb.AppendLine();
            }

            File.WriteAllText(saveDialog.FileName, sb.ToString(), Encoding.UTF8);

            TaskDialog.Show("Лог сохранён", "Лог успешно сохранён:\n" + saveDialog.FileName);
        }
    }
}
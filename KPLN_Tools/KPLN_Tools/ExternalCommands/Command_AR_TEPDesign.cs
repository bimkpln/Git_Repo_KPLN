using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Collections.Generic;

using View = Autodesk.Revit.DB.View;


namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_TEPDesign : IExternalCommand
    {
        UIApplication uiapp;

        ICollection<ElementId> viewportIds;

        int selectedCategory;
        int errorStatus = 0;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            if (!(activeView is ViewSheet viewSheet))
            {
                TaskDialog.Show("Предупреждение", "Выбранный вид не является листом.");
                return Result.Failed;
            }

            viewportIds = viewSheet.GetAllViewports();
            if (viewportIds.Count == 0)
            {
                TaskDialog.Show("Предупреждение", "На выбранном листе нет видов.");
                return Result.Failed;
            }


            ElementId textTypeId = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .FirstElementId();



            var dialogSelectCategory = new Forms.AR_TEPDesign_categorySelect();

            bool? dialogSelectCategoryResult = dialogSelectCategory.ShowDialog();
            selectedCategory = dialogSelectCategory.Result; 

            if (selectedCategory == 0)
            {
                TaskDialog.Show("Предупреждение", "Выбор категории отменён пользователем.");
                return Result.Failed;
            }           
            // Помещения
            else if (selectedCategory == 1)
            {
                HandlingCategoryRoom(doc, viewSheet); 

                if (errorStatus == 1)
                {
                    return Result.Failed;
                }
            }
            // Цветовые области (элементы узлов)
            else if (selectedCategory == 2)
            {
                if (errorStatus == 1)
                {
                    return Result.Failed;
                }
            }
            // Зоны
            else if (selectedCategory == 3)
            {
                if (errorStatus == 1)
                {
                    return Result.Failed;
                }
            }
            // Формы перекрытия
            else if (selectedCategory == 4)
            {
                if (errorStatus == 1)
                {
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }











        /// <summary>
        ///  "Помещения". Обработка категории "Помещения"
        /// </summary>
        public void HandlingCategoryRoom(Document doc, ViewSheet viewSheet)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            Dictionary<ElementId, Dictionary<ElementId, Room>> viewportsRoomsDict = new Dictionary<ElementId, Dictionary<ElementId, Room>>();

            foreach (ElementId vpId in viewportIds)
            {
                Viewport vp = doc.GetElement(vpId) as Viewport;
                if (vp == null) continue;

                View view = doc.GetElement(vp.ViewId) as View;
                if (view == null) continue;

                var rooms = new FilteredElementCollector(doc, view.Id)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .WhereElementIsNotElementType()
                            .OfType<Room>();

                if (!viewportsRoomsDict.TryGetValue(vpId, out var roomsDict))
                {
                    roomsDict = new Dictionary<ElementId, Room>();
                    viewportsRoomsDict[vpId] = roomsDict;
                }

                foreach (Room room in rooms)
                {
                    if (room == null || room.Area <= 0) continue;

                    if (!roomsDict.ContainsKey(room.Id))
                    {
                        roomsDict.Add(room.Id, room);
                    }
                }
            }

            List<ElementId> allRoomIds = viewportsRoomsDict
            .SelectMany(kvp => kvp.Value.Keys)
            .Distinct()
            .ToList();

            if (allRoomIds.Count == 0)
            {
                TaskDialog.Show("Предупреждение", "На данном листе не обнаружено виды, у которых имеются помещения.");
                return;
            }
            else
            {
                var categoryDialog = new Forms.AR_TEPDesign_paramNameSelect(doc, allRoomIds);
                bool? dialogResult = categoryDialog.ShowDialog(); 

                if (dialogResult != true)
                {
                    TaskDialog.Show("Предупреждение", "Выбор параметра отменён пользователем.");
                    errorStatus = 1;
                    return;
                }

                string selectedParamName = categoryDialog.SelectedParamName;
                string selectedEmptyLocation = categoryDialog.SelectedEmptyLocation;
                string selectedTableSortType = categoryDialog.SelectedTableSortType;
                System.Windows.Media.Color selectedColorEmptyColorScheme = categoryDialog.SelectedColorEmptyColorScheme;
                System.Windows.Media.Color selectedColorDummyСell = categoryDialog.SelectedColorDummyСell;
                bool SelectedELPriority = categoryDialog.SelectedELPriority;
                string colorBindingType = categoryDialog.SelectedColorBindingType;
                double selectedLightenFactor = categoryDialog.SelectedLightenFactor ?? 0.5;
                double selectedLightenFactorRow = categoryDialog.SelectedLightenFactorRow ?? 0.5;

                BuiltInCategory bic = BuiltInCategory.OST_Rooms;
                Dictionary<ElementId, Dictionary<string, Color>> parametersColor = getpParametersColor(doc, bic, viewportsRoomsDict, selectedParamName, selectedColorEmptyColorScheme, selectedLightenFactor);
                List<(ViewSchedule schedule, Color bgColor)> createdSchedules;

                using (var txCS = new Transaction(doc, "KPLN. Создание ТЭП-спецификаций"))
                {
                    txCS.Start();

                    // Удаление устаревших DraftingView
                    string viewName = $"ТЭП. Цветовая подложка - {viewSheet.SheetNumber}";
                    ViewDrafting existingDraftingView = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewDrafting))
                        .Cast<ViewDrafting>()
                        .FirstOrDefault(v => v.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase));

                    if (existingDraftingView != null) 
                    {
                        List<ElementId> toDelete = new List<ElementId>();

                        var relatedViewports = new FilteredElementCollector(doc)
                            .OfClass(typeof(Viewport))
                            .Cast<Viewport>()
                            .Where(vp => vp.ViewId == existingDraftingView.Id)
                            .ToList();

                        foreach (var vp in relatedViewports)
                        {
                            toDelete.Add(vp.Id);
                        }

                        toDelete.Add(existingDraftingView.Id);
                        if (toDelete.Count > 0)
                            doc.Delete(toDelete);
                    }
              
                    // Удаление устаревших спецификаций
                    string prefix = $"ТЭП_{viewSheet.SheetNumber} -";
                    FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule));
                    List<ViewSchedule> schedulesToDelete = scheduleCollector
                        .Cast<ViewSchedule>()
                        .Where(s => s.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    foreach (var schedule in schedulesToDelete)
                    {
                        doc.Delete(schedule.Id);
                    }

                    // Добавляем спецификацию через словарь
                    createdSchedules = CreateScheduleWithParam(doc, viewSheet, selectedParamName, parametersColor);

                    txCS.Commit();
                }

                using (var txAS = new Transaction(doc, "KPLN. Добавление ТЭП-спецификаций с оформлением на лист"))
                {
                    txAS.Start();
                    if (createdSchedules == null || !createdSchedules.Any())
                    {
                        TaskDialog.Show("Ошибка", "Не найдено значений для создания спецификаций.");
                        errorStatus = 1;
                        return;
                    }
                    else
                    {
                        List<(ViewSchedule schedule, Color bgColor)> sortedSchedules = SortSchedules(createdSchedules, selectedTableSortType);
                        addScheduleSheetInSheet(doc, viewSheet, sortedSchedules, selectedParamName, selectedEmptyLocation, selectedColorDummyСell, SelectedELPriority, colorBindingType, selectedLightenFactorRow);

                        if (errorStatus == 1)
                        {
                            txAS.RollBack();
                            return;
                        }
                    }
                    txAS.Commit();
                }
            }             
        }

        /// <summary>
        /// "Помещения". Составление словаря с параметром и сопутствующим ему цветом
        /// </summary>
        public Dictionary<ElementId, Dictionary<string, Color>> getpParametersColor(Document doc, BuiltInCategory bic, 
            Dictionary<ElementId, Dictionary<ElementId, Room>> viewportsRoomsDict, string selectedParamName, System.Windows.Media.Color selectedColorEmptyColorScheme, double selectedLightenFactor)
        {
            Dictionary<ElementId, Dictionary<string, Color>> result = new Dictionary<ElementId, Dictionary<string, Color>>();

            foreach (var kvp in viewportsRoomsDict)
            {
                ElementId viewportId = kvp.Key;
                Dictionary<ElementId, Room> roomsDict = kvp.Value;

                Dictionary<string, Color> valueColorMap = new Dictionary<string, Color>();

                foreach (var roomKvp in roomsDict)
                {
                    Room room = roomKvp.Value;
                    if (room == null) continue;

                    string paramValue = room.LookupParameter(selectedParamName)?.AsValueString() ?? room.LookupParameter(selectedParamName)?.AsString();

                    if (!string.IsNullOrEmpty(paramValue) && !valueColorMap.ContainsKey(paramValue))
                    {
                        Color color = GetColorFromColorScheme(doc, bic, viewportId, selectedParamName, paramValue);

                        if (color == null)
                        {
                            System.Windows.Media.Color wpfColor = selectedColorEmptyColorScheme;
                            color = new Autodesk.Revit.DB.Color(wpfColor.R, wpfColor.G, wpfColor.B);
                        }

                        if (selectedLightenFactor != 0.5)
                        {
                            color = AdjustColorBrightness(color, selectedLightenFactor);
                        }

                        valueColorMap[paramValue] = color;
                    }
                }

                result[viewportId] = valueColorMap;
            }

            return result;
        }



















        /// <summary>
        /// Цвет. Получение цвета параметра из цветовой схемы
        /// </summary>
        public Color GetColorFromColorScheme(Document doc, BuiltInCategory bic, ElementId elementId, string selectedParamName, string paramName)
        {
            var vp = doc.GetElement(elementId) as Viewport;
            if (vp == null)
            {
                return null;
            }

            var view = doc.GetElement(vp.ViewId) as View;
            if (view == null)
            {
                return null;
            }

            var cat = doc.Settings.Categories.get_Item(bic);
            if (cat == null)
            {
                return null;
            }

            List<ColorFillScheme> allSchemes = new FilteredElementCollector(doc)
                .OfClass(typeof(ColorFillScheme))
                .Cast<ColorFillScheme>()
                .Where(s => s.CategoryId == cat.Id)
                .ToList();

            ColorFillScheme scheme = allSchemes.FirstOrDefault(s =>
            {
                var pe = doc.GetElement(s.ParameterDefinition) as ParameterElement;
                string defName = pe != null
                    ? pe.Name
                    : LabelUtils.GetLabelFor((BuiltInParameter)s.ParameterDefinition.IntegerValue);
                return defName == selectedParamName;
            });

            if (scheme == null)
            {
                return null;
            }

            IList<ColorFillSchemeEntry> entries = scheme.GetEntries();
            ColorFillSchemeEntry match = entries.FirstOrDefault(e => e.GetStringValue() == paramName);

            if (match != null)
            {
                return match.Color;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Цвет. Изменения цвета ячеек
        /// </summary>
        Autodesk.Revit.DB.Color AdjustColorBrightness(Autodesk.Revit.DB.Color color, double factor)
        {
            byte Lerp(byte from, byte to, double tp) => (byte)(from + (to - from) * tp);

            if (factor == 0.5) return color;

            double t = (factor > 0.5) ? (factor - 0.5) * 2 : (0.5 - factor) * 2;
            byte r = Lerp(color.Red, factor > 0.5 ? (byte)255 : (byte)0, t);
            byte g = Lerp(color.Green, factor > 0.5 ? (byte)255 : (byte)0, t);
            byte b = Lerp(color.Blue, factor > 0.5 ? (byte)255 : (byte)0, t);

            return new Autodesk.Revit.DB.Color(r, g, b);
        }

        /// <summary>
        /// Словарь. Создание словаря со спецификациями и цветом
        /// <summary>
        public static List<(ViewSchedule schedule, Color bgColor)> CreateScheduleWithParam(Document doc, ViewSheet sheet, 
            string selectedParamName, Dictionary<ElementId, Dictionary<string, Color>> parametersColor)
        {
            const string sourceScheduleName = "Спецификация помещений";

            var sourceSched = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(vs => vs.Name.Equals(sourceScheduleName, StringComparison.OrdinalIgnoreCase));
            if (sourceSched == null)
            {
                TaskDialog.Show("Ошибка", $"Шаблон спецификации не найден");
                return null;
            }

            var uniqueValues = parametersColor
                .Values
                .SelectMany(inner => inner.Keys)
                .Distinct()
                .ToList();

            var createdSchedules = new List<(ViewSchedule, Color)>();
            string prefix = $"ТЭП_{sheet.SheetNumber} - {selectedParamName}";

            // Удаляем старые спецификации по заданному префиксу
            var oldIds = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => vs.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .Select(vs => vs.Id)
                .ToList();
            foreach (var id in oldIds)
                doc.Delete(id);

            // Создаём новые шаблоны
            foreach (var value in uniqueValues)
            {
                ElementId newId = sourceSched.Duplicate(ViewDuplicateOption.Duplicate);
                var sched = doc.GetElement(newId) as ViewSchedule;

                string invalidChars = new string(Path.GetInvalidFileNameChars()) + @"/:*?""<>|;";
                string safeValue = new string(value.Where(c => !invalidChars.Contains(c)).ToArray());
                sched.Name = $"{prefix} ({safeValue})";

                var def = sched.Definition;
                def.RemoveField(def.GetField(0).FieldId);

                var schedField = def.GetSchedulableFields()
                    .FirstOrDefault(f => f.GetName(doc)
                                           .Equals(selectedParamName, StringComparison.OrdinalIgnoreCase));
                var heightField = def.GetSchedulableFields()
                    .FirstOrDefault(f => f.GetName(doc)
                           .Equals("Высота строки", StringComparison.OrdinalIgnoreCase));

                if (schedField == null)
                {
                    TaskDialog.Show("Ошибка", $"Параметр \"{selectedParamName}\" не найден в списке полей.");
                    return null;
                }            
                def.InsertField(schedField, 0);
                ScheduleFieldId fieldId = def.GetFieldId(0);

                def.AddFilter(new ScheduleFilter(fieldId, ScheduleFilterType.Equal,value));

                Color bgColor = parametersColor
                        .SelectMany(kvp => kvp.Value)
                        .First(pair => pair.Key.Equals(value, StringComparison.OrdinalIgnoreCase))
                        .Value;

                    createdSchedules.Add((sched, bgColor));                           
            }

            return createdSchedules;
        }

        /// <summary>
        /// Словарь. Сортировка спецификации по параметру
        /// </summary>
        public static List<(ViewSchedule schedule, Color bgColor)> SortSchedules(
            List<(ViewSchedule schedule, Color bgColor)> schedules,
            string sortType)
        {
            Func<ViewSchedule, string> getFirstCellValue = schedule =>
            {
                var section = schedule.GetTableData().GetSectionData(SectionType.Body);
                return schedule.GetCellText(SectionType.Body, section.FirstRowNumber, section.FirstColumnNumber);
            };

            Func<ViewSchedule, string> getSecondCellRaw = schedule =>
            {
                var section = schedule.GetTableData().GetSectionData(SectionType.Body);
                return schedule.GetCellText(SectionType.Body, section.FirstRowNumber, section.FirstColumnNumber + 1);
            };

            string CleanCellText(string input)
            {
                if (string.IsNullOrWhiteSpace(input)) return "";
                input = input.ToLowerInvariant();
                input = input.Replace("м²", "");
                input = new string(input.Where(c => !char.IsWhiteSpace(c)).ToArray());
                input = input.Replace(",", ".");
                return input;
            }

            Func<ViewSchedule, double> getSecondCellValueAsNumber = schedule =>
            {
                string raw = getSecondCellRaw(schedule);
                string cleaned = CleanCellText(raw);

                return double.TryParse(cleaned, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double result)
                    ? result
                    : double.MinValue;
            };

            switch (sortType)
            {
                case "Значение параметра. А-Я":
                    return schedules
                        .OrderBy(item => getFirstCellValue(item.schedule))
                        .ThenBy(item => getSecondCellValueAsNumber(item.schedule))
                        .ToList();

                case "Значение параметра. Я-А":
                    return schedules
                        .OrderByDescending(item => getFirstCellValue(item.schedule))
                        .ThenByDescending(item => getSecondCellValueAsNumber(item.schedule))
                        .ToList();

                default:
                    return schedules;
            }
        }

        /// <summary>
        /// Лист. Добавление спецификаций на лист
        /// </summary>
        public void addScheduleSheetInSheet(Document doc, ViewSheet viewSheet, List<(ViewSchedule schedule, Color bgColor)> createdSchedules,
           string selectedParamName, string selectedEmptyLocation, System.Windows.Media.Color selectedColorDummyСell, bool SelectedELPriority, string colorBindingType, double selectedLightenFactorRow)
        {
            // Размеры ячеек
            const double mmToFeet = 0.00328084;
            double startX = 20 * mmToFeet;
            double startY = 51.4 * mmToFeet;
            double heightFrame = 395 * mmToFeet;
            double rowStep = 10 * mmToFeet;

            int count = createdSchedules.Count;
            int colCount = (count % 2 == 0) ? count / 2 : (count + 1) / 2;
            double widthPerSchedule = heightFrame / colCount;

            List<(int row, int col, (ViewSchedule schedule, Color color)? data)> layout =
                new List<(int row, int col, (ViewSchedule schedule, Color color)? data)>();

            int currentIndex = 0;
            int totalSlots = colCount * 2;

            Category linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            GraphicsStyle invisibleLineStyle = linesCategory.SubCategories
                .Cast<Category>()
                .Select(sub => sub.GetGraphicsStyle(GraphicsStyleType.Projection))
                .FirstOrDefault(style => style != null && style.Name == "<Невидимые линии>");

            if (invisibleLineStyle == null)
            {
                TaskDialog.Show("Ошибка", "Не найден стиль '<Невидимые линии>'");
                return;
            }

            for (int i = 0; i < totalSlots; i++)
            {
                int row = i / colCount;
                int col = i % colCount;

                bool insertEmpty = false;

                if (selectedEmptyLocation == "Сверху слева" && i == 0)
                    insertEmpty = true;
                else if (selectedEmptyLocation == "Сверху справа" && i == colCount - 1)
                    insertEmpty = true;
                else if (selectedEmptyLocation == "Снизу слева" && i == colCount)
                    insertEmpty = true;
                else if (selectedEmptyLocation == "Снизу справа" && i == totalSlots - 1)
                    insertEmpty = true;

                if (insertEmpty || currentIndex >= createdSchedules.Count)
                {
                    layout.Add((row, col, null));
                }
                else
                {
                    layout.Add((row, col, createdSchedules[currentIndex]));
                    currentIndex++;
                }
            }

            // Перекрашивание ячеек (уникальный первый столбец)
            if (colorBindingType == "Первый ряд уникальный")
            {
                for (int i = 0; i < layout.Count; i++)
                {
                    var (row, col, data) = layout[i];
                    if (row != 1 || data == null)
                        continue;


                    var topCell = layout.FirstOrDefault(x => x.row == 0 && x.col == col);
                    System.Windows.Media.Color baseColor;

                    if (topCell.data != null)
                    {
                        var c = topCell.data.Value.color;
                        baseColor = System.Windows.Media.Color.FromRgb(c.Red, c.Green, c.Blue);
                    }
                    else
                    {
                        baseColor = selectedColorDummyСell;
                    }

                    byte R = (byte)Math.Min(255, baseColor.R + (255 - baseColor.R) * selectedLightenFactorRow);
                    byte G = (byte)Math.Min(255, baseColor.G + (255 - baseColor.G) * selectedLightenFactorRow);
                    byte B = (byte)Math.Min(255, baseColor.B + (255 - baseColor.B) * selectedLightenFactorRow);
                    var lightenedColor = new Autodesk.Revit.DB.Color(R, G, B);

                    layout[i] = (row, col, (data.Value.schedule, lightenedColor));
                }
            }

            // ViewDrafting
            string viewName = $"ТЭП. Цветовая подложка - {viewSheet.SheetNumber}";
            ViewDrafting draftingView = ViewDrafting.Create(doc, doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeDrafting));

            draftingView.Name = viewName;
            draftingView.Scale = 1;

            ElementId solidFillPatternId = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill)?.Id ?? ElementId.InvalidElementId;

            if (solidFillPatternId == ElementId.InvalidElementId)
            {
                TaskDialog.Show("Ошибка", "Не найден шаблон заливки");
                return;
            }

            FilledRegionType baseRegionType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (baseRegionType == null)
            {
                TaskDialog.Show("Ошибка", "Не найден тип заливки.");
                return;
            }

            foreach (var (row, col, data) in layout)
            {
                // Заглушка
                if (data == null)
                {
                    double x0 = startX + col * widthPerSchedule;
                    double y0 = startY - row * rowStep;
                    double x1 = x0 + widthPerSchedule;
                    double y1 = y0 - rowStep;

                    List<Curve> curves = new List<Curve>
                    {
                        Line.CreateBound(new XYZ(x0, y0, 0), new XYZ(x1, y0, 0)),
                        Line.CreateBound(new XYZ(x1, y0, 0), new XYZ(x1, y1, 0)),
                        Line.CreateBound(new XYZ(x1, y1, 0), new XYZ(x0, y1, 0)),
                        Line.CreateBound(new XYZ(x0, y1, 0), new XYZ(x0, y0, 0))
                    };

                    CurveLoop loop = CurveLoop.Create(curves);

                    string typeName = $"ТЭП. {viewSheet.SheetNumber} ({selectedParamName} - ЗАГЛУШКА)";

                    var existing = new FilteredElementCollector(doc)
                        .OfClass(typeof(FilledRegionType))
                        .Cast<FilledRegionType>()
                        .FirstOrDefault(t => t.Name == typeName);

                    if (existing != null)
                    {
                        doc.Delete(existing.Id);
                    }

                    FilledRegionType newType = baseRegionType.Duplicate(typeName) as FilledRegionType;
                    newType.ForegroundPatternId = solidFillPatternId;

                    if (SelectedELPriority)
                    {
                        newType.ForegroundPatternColor = new Autodesk.Revit.DB.Color(selectedColorDummyСell.R, selectedColorDummyСell.G, selectedColorDummyСell.B);
                    }
                    else
                    {
                        Autodesk.Revit.DB.Color finalColor;
                        if (row == 1 && colorBindingType == "Первый ряд уникальный")
                        {
                            // Ищем ячейку сверху
                            var topCell = layout.FirstOrDefault(x => x.row == 0 && x.col == col);

                            System.Windows.Media.Color baseColor;

                            if (topCell.data != null)
                            {
                                var c = topCell.data.Value.color;
                                baseColor = System.Windows.Media.Color.FromRgb(c.Red, c.Green, c.Blue);
                            }
                            else
                            {
                                baseColor = selectedColorDummyСell;
                            }

                            byte R = (byte)Math.Min(255, baseColor.R + (255 - baseColor.R) * selectedLightenFactorRow);
                            byte G = (byte)Math.Min(255, baseColor.G + (255 - baseColor.G) * selectedLightenFactorRow);
                            byte B = (byte)Math.Min(255, baseColor.B + (255 - baseColor.B) * selectedLightenFactorRow);

                            finalColor = new Autodesk.Revit.DB.Color(R, G, B);
                        }
                        else
                        {
                            finalColor = new Autodesk.Revit.DB.Color(
                                selectedColorDummyСell.R,
                                selectedColorDummyСell.G,
                                selectedColorDummyСell.B
                            );
                        }
                        newType.ForegroundPatternColor = finalColor;
                    }                  

                    newType.BackgroundPatternId = ElementId.InvalidElementId;

                    ElementId typeId = newType.Id;

                    FilledRegion region = FilledRegion.Create(doc, typeId, draftingView.Id, new List<CurveLoop> { loop });
                    var curveIds = region.GetDependentElements(new ElementClassFilter(typeof(CurveElement)));

                    foreach (ElementId curveId in curveIds)
                    {
                        CurveElement curveElement = doc.GetElement(curveId) as CurveElement;
                        if (curveElement is DetailCurve detailCurve)
                        {
                            detailCurve.LineStyle = invisibleLineStyle;
                        }
                    }
                }
                // Всё остальное
                else
                {
                    var (schedule, color) = data.Value;

                    double x0 = startX + col * widthPerSchedule;
                    double y0 = startY - row * rowStep;
                    double x1 = x0 + widthPerSchedule;
                    double y1 = y0 - rowStep;

                    List<Curve> curves = new List<Curve>
                    {
                        Line.CreateBound(new XYZ(x0, y0, 0), new XYZ(x1, y0, 0)),
                        Line.CreateBound(new XYZ(x1, y0, 0), new XYZ(x1, y1, 0)),
                        Line.CreateBound(new XYZ(x1, y1, 0), new XYZ(x0, y1, 0)),
                        Line.CreateBound(new XYZ(x0, y1, 0), new XYZ(x0, y0, 0))
                    };

                    CurveLoop loop = CurveLoop.Create(curves);

                    TableData tableData = schedule.GetTableData();
                    TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);
                    string valueTab00 = schedule.GetCellText(SectionType.Body, 0, 0);

                    string typeName = $"ТЭП. {viewSheet.SheetNumber} ({selectedParamName} - {valueTab00})";

                    var existing = new FilteredElementCollector(doc)
                        .OfClass(typeof(FilledRegionType))
                        .Cast<FilledRegionType>()
                        .FirstOrDefault(t => t.Name == typeName);

                    if (existing != null)
                    {
                        doc.Delete(existing.Id);
                    }

                    FilledRegionType newType = baseRegionType.Duplicate(typeName) as FilledRegionType;
                    newType.ForegroundPatternId = solidFillPatternId;
                    newType.ForegroundPatternColor = color;
                    newType.BackgroundPatternId = ElementId.InvalidElementId;

                    ElementId typeId = newType.Id;

                    FilledRegion region = FilledRegion.Create(doc, typeId, draftingView.Id, new List<CurveLoop> { loop });
                    var curveIds = region.GetDependentElements(new ElementClassFilter(typeof(CurveElement)));

                    foreach (ElementId curveId in curveIds)
                    {
                        CurveElement curveElement = doc.GetElement(curveId) as CurveElement;
                        if (curveElement is DetailCurve detailCurve)
                        {
                            detailCurve.LineStyle = invisibleLineStyle;
                        }
                    }
                }
            }

            doc.Regenerate();

            XYZ insertPoint = new XYZ(217.5 * mmToFeet, 41.4 * mmToFeet, 0); // Точка середины марки
            Viewport vp = Viewport.Create(doc, viewSheet.Id, draftingView.Id, insertPoint);

            Parameter titleOnSheetParam = vp.LookupParameter("Заголовок на листе");

            if (titleOnSheetParam != null && !titleOnSheetParam.IsReadOnly)
            {
                titleOnSheetParam.Set("\u200B"); // невидимый символ (Zero-Width Space)
            }

            // Спецификации
            foreach (var (row, col, data) in layout)
            {
                if (data == null)
                    continue;

                var (schedule, color) = data.Value;

                // Установка ширины столбцов
                TableData tableData = schedule.GetTableData();
                TableSectionData body = tableData.GetSectionData(SectionType.Body);

                if (body.NumberOfColumns > 2)
                {
                    body.SetColumnWidth(0, widthPerSchedule / 2);
                    body.SetColumnWidth(1, widthPerSchedule / 2);
                    body.SetColumnWidth(2, 0.0000000000001);
                }

                double offsetX = col * widthPerSchedule;
                double offsetY = -row * rowStep;

                XYZ point = new XYZ(startX + offsetX, startY + offsetY, 0);
                ScheduleSheetInstance.Create(doc, viewSheet.Id, schedule.Id, point);
            }
        }
    }
}
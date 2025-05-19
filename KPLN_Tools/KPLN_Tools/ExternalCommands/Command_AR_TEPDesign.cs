using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using View = Autodesk.Revit.DB.View;
using System.Text;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_TEPDesign : IExternalCommand
    {        
        ICollection<ElementId> viewportIds;

        int selectedCategory;
        private const string CellFillTypeName = "KPLN_CellFill";
        int errorStatus = 0;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
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

            double fontSize = 0.5;
            TextNoteType textNoteType = doc.GetElement(textTypeId) as TextNoteType;
            if (textNoteType != null)
            {
                double internalSize = textNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble();

#if (Debug2020 || Revit2020)
    fontSize = UnitUtils.ConvertFromInternalUnits(internalSize, DisplayUnitType.DUT_MILLIMETERS);
#endif
#if (Debug2023 || Revit2023)
                fontSize = UnitUtils.ConvertFromInternalUnits(internalSize, UnitTypeId.Millimeters);
#endif
            }

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
                HandlingCategoryRoom(doc, viewSheet, fontSize); 

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













        // Обработка категории "Помещения"
        public void HandlingCategoryRoom(Document doc, ViewSheet viewSheet, double fontSize)
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
                var categoryDialog = new Forms.AR_TEPDesign_paramNameSelect(doc, allRoomIds, fontSize);
                bool? dialogResult = categoryDialog.ShowDialog(); 

                if (dialogResult != true)
                {
                    TaskDialog.Show("Предупреждение", "Выбор параметра отменён пользователем.");
                    errorStatus = 1;
                    return;
                }

                string selectedParamName = categoryDialog.SelectedParamName;
                double tableHeightInternal = categoryDialog.SelectedTableHeight ?? 0.01;
                double fontSizeInternal = categoryDialog.SelectedFontSize ?? 0.01;
                double selectedLightenFactor = categoryDialog.SelectedLightenFactor ?? 0.01;
                double selectedDarkenFactor = categoryDialog.SelectedDarkenFactor ?? 0.01;

                if (tableHeightInternal == 0.01 || fontSizeInternal == 0.01)
                {
                    TaskDialog.Show("Ошибка", "Не удалось обработать значения высоты таблицы и размера шрифта.");
                    errorStatus = 1;
                    return;
                }

                Dictionary<ElementId, Dictionary<string, double>> areaByParamValue = GetAreaSummaryByParam(doc, viewportsRoomsDict, selectedParamName);

                if (areaByParamValue == null || areaByParamValue.Count == 0)
                {
                    TaskDialog.Show("Ошибка", $"Значения параметра {selectedParamName} пусты или элементы не содержат площадь");
                    errorStatus = 1;
                    return; 
                }
                else
                {
                    BuiltInCategory bic = BuiltInCategory.OST_Rooms;
                    CreateDraftingView(doc, bic, viewSheet, selectedParamName, areaByParamValue, 
                        tableHeightInternal, fontSizeInternal, selectedLightenFactor, selectedDarkenFactor);
                }                     
            }             
        }

        // Словарь: ID-плана - <ID-комнаты, комната>
        public Dictionary<ElementId, Dictionary<string, double>> GetAreaSummaryByParam(Document doc, 
            Dictionary<ElementId, Dictionary<ElementId, Room>> viewportsRoomsDict, string selectedParamName)
        {
            var result = new Dictionary<ElementId, Dictionary<string, double>>();

            foreach (var vpEntry in viewportsRoomsDict)
            {
                ElementId vpId = vpEntry.Key;
                Dictionary<ElementId, Room> roomsDict = vpEntry.Value;

                Dictionary<string, double> areaByParamValue = new Dictionary<string, double>();

                foreach (var roomEntry in roomsDict)
                {
                    Room room = roomEntry.Value;
                    if (room == null) continue;

                    string groupingValue = room.LookupParameter(selectedParamName)?.AsString() ?? "—";

                    double area = 0;

                    Parameter areaParam = room.LookupParameter("Площадь");

#if (Debug2020 || Revit2020)
            if (areaParam != null && areaParam.StorageType == StorageType.Double)
            {
                area = UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
            }
            else
            {
                area = UnitUtils.ConvertFromInternalUnits(room.Area, DisplayUnitType.DUT_SQUARE_METERS);
            }
#endif
#if (Debug2023 || Revit2023)
                    if (areaParam != null && areaParam.StorageType == StorageType.Double)
                    {
                        area = UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);
                    }
                    else
                    {
                        area = UnitUtils.ConvertFromInternalUnits(room.Area, UnitTypeId.SquareMeters);
                    }
#endif
                    if (areaByParamValue.ContainsKey(groupingValue))
                    {
                        areaByParamValue[groupingValue] += area;
                    }
                    else
                    {
                        areaByParamValue[groupingValue] = area;
                    }
                }

                areaByParamValue.Remove("—");

                areaByParamValue = areaByParamValue
                    .OrderBy(entry => entry.Key)
                    .ToDictionary(entry => entry.Key, entry => entry.Value);

                if (areaByParamValue.Count > 0)
                {
                    result[vpId] = areaByParamValue;
                }
            }

            return result;
        }















        // Создание и добавление DraftingView
        public void CreateDraftingView(Document doc, BuiltInCategory bic, ViewSheet sheet, 
            string selectedParamName, Dictionary<ElementId, Dictionary<string, double>> areaByParamValue, 
            double tableHeightInternal, double fontSizeInternal, double selectedLightenFactor, double selectedDarkenFactor)
        {
            if (doc == null || sheet == null || areaByParamValue == null)
            {
                TaskDialog.Show("Ошибка", "Входные данные некорректны.");
                return;
            }

            ViewDrafting tableView;

            using (Transaction t = new Transaction(doc, "KPLN. Создание ТЭП-таблицы"))
            {
                t.Start();

                // Создание чертёжного вида
                string viewTableName = $"ТЭП. {sheet.SheetNumber}";

                ViewDrafting existingView = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>()
                    .FirstOrDefault(v => v.Name == viewTableName);

                if (existingView != null)
                {
                    Viewport existingVP = new FilteredElementCollector(doc)
                        .OfClass(typeof(Viewport))
                        .Cast<Viewport>()
                        .FirstOrDefault(vp => vp.ViewId == existingView.Id);

                    if (existingVP != null)
                        doc.Delete(existingVP.Id);

                    doc.Delete(existingView.Id);
                }

                tableView = ViewDrafting.Create(doc, doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeDrafting));
                tableView.Name = viewTableName;

                // Получение размера основной надписи
                FamilyInstance titleBlock = new FilteredElementCollector(doc, sheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .FirstOrDefault();

                BoundingBoxXYZ bboxTitleBlock = titleBlock?.get_BoundingBox(sheet);
                if (bboxTitleBlock == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось определить ширину основной надписи.");
                    return;
                }

                double width = (bboxTitleBlock.Max.X - bboxTitleBlock.Min.X) * 10; // Коэффициени "Маштаб вида"

                // Создание словаря-таблицы
                Dictionary<int, Dictionary<int, (string paramValue, double area, Color color)>> tableDict = BuildTable(doc, bic, selectedParamName, areaByParamValue, selectedLightenFactor, selectedDarkenFactor);

                int columns = tableDict.Any() ? tableDict.First().Value.Count : 0;
                double cellWidth = width / columns;
                double cellHeight = tableHeightInternal;

                var categoriesTableDicts = Enumerable.Range(0, columns).Select(colIndex => tableDict.Select(row => row.Value[colIndex].paramValue) 
                          .First(pv => !string.IsNullOrEmpty(pv))).ToList();

                // Формирование текста внутри ячейки
                ElementId textTypeId = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .FirstElementId();

                TextNoteType textNoteType = doc.GetElement(textTypeId) as TextNoteType;
                textNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(fontSizeInternal);

                var optsLeft = new TextNoteOptions
                {
                    TypeId = textTypeId,
                    HorizontalAlignment = HorizontalTextAlignment.Left,
                    VerticalAlignment = VerticalTextAlignment.Middle
                };

                var optsRight = new TextNoteOptions
                {
                    TypeId = textTypeId,
                    HorizontalAlignment = HorizontalTextAlignment.Right,
                    VerticalAlignment = VerticalTextAlignment.Middle
                };

                double leftPadding = 15.0 / 304.8;
                double rightPadding = 15.0 / 304.8;

                // Добавление данных в таблицу
                int rowIndex = 0;
                foreach (var kvp in tableDict.OrderBy(r => r.Key))
                {
                    for (int col = 0; col < columns; col++)
                    {
                        string categoryKVP = categoriesTableDicts[col];      
                        var cell = kvp.Value[col];
                        Color fillColor = cell.color;

                        double x0 = col * cellWidth;
                        double y0 = -rowIndex * cellHeight;
                        double x1 = x0 + cellWidth;
                        double y1 = y0 - cellHeight;

                        // заливаем задний фон ячейки из cell.color
                        FillCellBackground(doc, tableView, fillColor, x0, y0, x1, y1);

                        if (!string.IsNullOrEmpty(cell.paramValue))
                        {
                            double centerY = (y0 + y1) / 2;
                            string valueKVP = $"{cell.area:F2} м²";

                            XYZ posLeft = new XYZ(x0 + leftPadding, centerY, 0);
                            XYZ posRight = new XYZ(x1 - rightPadding, centerY, 0);

                            var noteLeft = TextNote.Create(doc, tableView.Id, posLeft,
                                                            categoryKVP, optsLeft);
                            var noteRight = TextNote.Create(doc, tableView.Id, posRight,
                                                            valueKVP, optsRight);

                            double cellDivisionFactor = 1 + Math.Pow(columns, 0.6) / 5;
                            noteLeft.Width = ((cellWidth - leftPadding) / 10) / cellDivisionFactor;
                        }
                    }

                    rowIndex++;
                }

                t.Commit();
            }

            PlaceDraftingViewOnSheet(doc, sheet, tableView);
        }







        // Таблица-словарь. Формирование
        public Dictionary<int, Dictionary<int, (string paramValue, double area, Color color)>> BuildTable(Document doc, BuiltInCategory bic, string selectedParamName,
            Dictionary<ElementId, Dictionary<string, double>> areaByParamValue, double selectedLightenFactor, double selectedDarkenFactor)
        {
            var table = new Dictionary<int, Dictionary<int, (string paramValue, double area, Color color)>>();

            List<string> allParamValuesName = areaByParamValue.Values
                .SelectMany(dict => dict.Keys)
                .Distinct()
                .OrderBy(val => val)
                .ToList();

            List<ElementId> elementIdsList = areaByParamValue.Keys.ToList();

            double darkenFactor = selectedDarkenFactor;
            double lightenFactor = selectedLightenFactor;

            // Функции затемнения
            Color Darken(Color color, double factor)
            {
                byte DarkenComponent(byte comp)
                {
                    double val = comp * (1 - factor);
                    if (val < 0) val = 0;
                    if (val > 255) val = 255;
                    return (byte)val;
                }
                return new Color(
                    DarkenComponent(color.Red),
                    DarkenComponent(color.Green),
                    DarkenComponent(color.Blue));
            }

            // Функция осветления
            Color Lighten(Color color, double factor)
            {
                byte LightenComponent(byte comp)
                {
                    double val = comp + (255 - comp) * factor;
                    if (val < 0) val = 0;
                    if (val > 255) val = 255;
                    return (byte)val;
                }
                return new Color(
                    LightenComponent(color.Red),
                    LightenComponent(color.Green),
                    LightenComponent(color.Blue));
            }
          
            // Строки таблицы
            for (int rowIndex = 0; rowIndex < elementIdsList.Count; rowIndex++)
            {
                Dictionary<string, double> rowValues = areaByParamValue[elementIdsList[rowIndex]];
                Dictionary<int, (string, double, Color)> rowDict = new Dictionary<int, (string, double, Color)>();

                // Столбцы таблицы
                for (int colIndex = 0; colIndex < allParamValuesName.Count; colIndex++)
                {
                    string paramName = allParamValuesName[colIndex];

                    if (rowValues.TryGetValue(paramName, out double area))
                    {

                        List<int> rowsWithValue = elementIdsList
                            .Select((id, idx) => new { id, idx })
                            .Where(x => areaByParamValue[x.id].ContainsKey(paramName))
                            .Select(x => x.idx)
                            .OrderBy(idx => idx)
                            .ToList();

                        Color cellColor = GetColorFromColorScheme(doc, bic, elementIdsList[rowIndex], selectedParamName, paramName);

                        rowDict[colIndex] = (paramName, area, cellColor);
                    }
                    else
                    {
                        rowDict[colIndex] = (string.Empty, 0.0, null);
                    }
                }

                table[rowIndex] = rowDict;
            }

            int rowCount = elementIdsList.Count;
            int colCount = allParamValuesName.Count;

            // Обработка пустых цветов
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < colCount; colIndex++)
                {
                    var cell = table[rowIndex][colIndex];
                    if (cell.color == null)
                    {
                        Color newColor;

                        if (rowIndex > 0 && table[rowIndex - 1][colIndex].color != null)
                        {
                            newColor = Lighten(
                                table[rowIndex - 1][colIndex].color,
                                lightenFactor);
                        }
                        else if (rowIndex < rowCount - 1 && table[rowIndex + 1][colIndex].color != null)
                        {
                            newColor = Darken(
                                table[rowIndex + 1][colIndex].color,
                                darkenFactor);
                        }
                        else
                        {
                            newColor = new Color(255, 255, 255);
                        }

                        table[rowIndex][colIndex] =
                            (cell.paramValue, cell.area, newColor);
                    }
                }
            }

            return table;
        }

        // Получение цвета из схемы
        public Color GetColorFromColorScheme(Document doc, BuiltInCategory bic, ElementId elementId, string selectedParamName, string paramName)
        {
            var vp = doc.GetElement(elementId) as Viewport;
            if (vp == null)
            {
                TaskDialog.Show("Ошибка", "Первый ID не является Viewport.");
                return new Autodesk.Revit.DB.Color(255, 255, 255);
            }
            var view = doc.GetElement(vp.ViewId) as View;
            if (view == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось получить View из Viewport.");
                return new Autodesk.Revit.DB.Color(255, 255, 255);
            }

            var cat = doc.Settings.Categories.get_Item(bic);
            if (cat == null)
            {
                TaskDialog.Show("Ошибка", $"Категория {bic} не найдена.");
                return new Autodesk.Revit.DB.Color(255, 255, 255);
            }

            var allSchemes = new FilteredElementCollector(doc)
                .OfClass(typeof(ColorFillScheme))
                .Cast<ColorFillScheme>()
                .Where(s => s.CategoryId == cat.Id)
                .ToList();

            var scheme = allSchemes.FirstOrDefault(s =>
            {
                var pe = doc.GetElement(s.ParameterDefinition) as ParameterElement;
                string defName = pe != null
                    ? pe.Name
                    : LabelUtils.GetLabelFor((BuiltInParameter)s.ParameterDefinition.IntegerValue);
                return defName == selectedParamName;
            });

            if (scheme == null)
            {
                TaskDialog.Show("Ошибка",
                    $"Не найдена цветовая схема для категории «{cat.Name}», привязанная к параметру «{selectedParamName}».");
                return new Autodesk.Revit.DB.Color(255, 255, 255);
            }

            var entries = scheme.GetEntries();
            var match = entries.FirstOrDefault(e => e.GetStringValue() == paramName);

            if (match != null)
            {
                return match.Color;
            }
            else
            {
                TaskDialog.Show("Ошибка",
                    $"В ColorFillScheme «{scheme.Name}» нет записи с GetStringValue() = «{paramName}».");
                return new Autodesk.Revit.DB.Color(255, 255, 255);
            }
        }

        // DraftingView. Расскраска ячеек
        private static void FillCellBackground(Document doc, View tableView, Color fillColor, double x0, double y0, double x1, double y1)
        {
            var fillType = GetOrCreateCellFillType(doc);

            if (fillColor == null) return;      
            fillType.ForegroundPatternColor = fillColor;
            fillType.BackgroundPatternColor = fillColor;
            
            var loop = new CurveLoop();
            loop.Append(Line.CreateBound(new XYZ(x0, y0, 0), new XYZ(x1, y0, 0)));
            loop.Append(Line.CreateBound(new XYZ(x1, y0, 0), new XYZ(x1, y1, 0)));
            loop.Append(Line.CreateBound(new XYZ(x1, y1, 0), new XYZ(x0, y1, 0)));
            loop.Append(Line.CreateBound(new XYZ(x0, y1, 0), new XYZ(x0, y0, 0)));

            var region = FilledRegion.Create(
                doc,
                fillType.Id,
                tableView.Id,
                new List<CurveLoop> { loop });

            var ogs = new OverrideGraphicSettings()
                .SetProjectionLineColor(fillColor)
                .SetProjectionLineWeight(1)
                .SetCutLineColor(fillColor)
                .SetCutLineWeight(1);

            tableView.SetElementOverrides(region.Id, ogs);
        }

        // DraftingView. Вспомогательный метод поиска стилей
        private static FilledRegionType GetOrCreateCellFillType(Document doc)
        {
            const string typeName = "ThisTypeDoesNotExist"; // Это костыль при формировании типа

            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(frt => frt.Name == typeName);
            if (existing != null)
                return existing;

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault()
                ?? throw new InvalidOperationException("Нет ни одного FilledRegionType для дублирования.");

            var newType = baseType.Duplicate(typeName) as FilledRegionType
                          ?? throw new InvalidOperationException("Не удалось дублировать тип.");

            var solidId = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill)
                ?.Id ?? ElementId.InvalidElementId;

            if (solidId != ElementId.InvalidElementId)
            {
                newType.ForegroundPatternId = solidId;
                newType.BackgroundPatternId = solidId;
            }

            return newType;
        }

















        // Добавление DraftingView на страницу
        public void PlaceDraftingViewOnSheet(Document doc, ViewSheet sheet, ViewDrafting view)
        {
            using (Transaction t = new Transaction(doc, "KPLN. Размещение вида в рамке"))
            {
                t.Start();

                var existingVP = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .FirstOrDefault(vpO => vpO.ViewId == view.Id);

                if (existingVP != null)
                    doc.Delete(existingVP.Id);

                var titleBlock = new FilteredElementCollector(doc, sheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .FirstElement();

                if (titleBlock == null)
                {
                    TaskDialog.Show("Ошибка", "Не найдена рамка на листе, чтобы разместить DraftingView.");
                    return;
                }

                BoundingBoxXYZ bbox = titleBlock.get_BoundingBox(sheet);
                if (bbox == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось получить границы рамки.");
                    return;
                }

                // Добавим небольшой отступ от рамки
                double offset = 0.2;
                XYZ bottomLeft = new XYZ(bbox.Min.X + offset, bbox.Min.Y + offset, 0);

                // Размещаем DraftingView
                Viewport vp = Viewport.Create(doc, sheet.Id, view.Id, bottomLeft);
                if (vp == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось разместить таблицу на листе.");
                    return;
                }

                view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)?.Set("\u200B");

                t.Commit();
            }
        }
    }
}
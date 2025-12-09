using Autodesk.Revit.DB;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    public sealed class ScheduleEntity
    {
        public ViewSchedule SE_ViewSchedule { get; set; }
        public List<ScheduleFieldId> SE_HiddenFieldIds { get; set; } = new List<ScheduleFieldId>();
    }

    internal sealed class GroupFieldInfo
    {
        /// <summary>
        /// Ссылка на поле
        /// </summary>
        public ScheduleField Field { get; set; }
        
        /// <summary>
        /// Индекс столбца
        /// </summary>
        public int ColumnIndex { get; set; }
    }

    internal static class ScheduleHelper
    {
        public static ScheduleFormM ReadSchedule(ScheduleEntity se)
        {
            TableData tableData = se.SE_ViewSchedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);
            ScheduleDefinition def = se.SE_ViewSchedule.Definition;

            int rows = body.NumberOfRows;
            int cols = body.NumberOfColumns;
            int allCols = def.GetFieldCount();

            var model = new ScheduleFormM
            {
                SE_Schedule = se
            };


            // 1) Заголовки столбцов
            for (int i = 0; i < allCols; i++)
            {
                ScheduleField field = def.GetField(i);
                if (field.IsHidden)
                    continue;

                string headerText = field.GetName();
                if (string.IsNullOrEmpty(headerText))
                    headerText = $"Col {i + 1}";

                model.ColumnHeaders.Add(new CellData(headerText, false));
            }

            // 2) Ряды таблицы
            for (int row = 0; row < rows; row++)
            {
                var rowData = new List<CellData>();
                for (int col = 0; col < cols; col++)
                {
                    CellType cellType = body.GetCellType(row, col);
                    bool isElementRow = (cellType == CellType.Parameter || cellType == CellType.CombinedParameter);
                    string cellText = se.SE_ViewSchedule.GetCellText(SectionType.Body, row, col);
                    cellText = DoubleWithoutCulture(cellText);

                    rowData.Add(new CellData(cellText, isElementRow));
                }

                model.Rows.Add(rowData);
            }

            return model;
        }

        /// <summary>
        /// Строим список группиорванных по полям элементов со спеки
        /// </summary>
        public static List<GroupFieldInfo> GetGroupFields(ViewSchedule vs)
        {
            var result = new List<GroupFieldInfo>();

            ScheduleDefinition def = vs.Definition;
            for (int i = 0; i < def.GetSortGroupFieldCount(); i++)
            {
                var sg = def.GetSortGroupField(i);
                ScheduleField field = def.GetField(sg.FieldId);

                result.Add(new GroupFieldInfo
                {
                    Field = field,
                    ColumnIndex = field.FieldIndex
                });
            }

            return result;
        }

        /// <summary>
        /// Вяртае значэнне поля для элемента, як максімальна падобнае да таго, што ў спецыфікацыі.
        /// </summary>
        public static string GetFieldValueForElement(Document doc, Element el, ScheduleField field)
        {
            Parameter p = null;

#if Revit2020 || Debug2020 || Revit2023 || Debug2023
            BuiltInParameter parBIC = (BuiltInParameter)field.ParameterId.IntegerValue;
#else
            BuiltInParameter parBIC = (BuiltInParameter)field.ParameterId.Value;
#endif

                // Параметр экз.
                if (field.ParameterId != ElementId.InvalidElementId)
                p = el.get_Parameter(parBIC);
            else
                p = el.LookupParameter(field.GetName());

            // Параметр типа
            if (!p.HasValue )
            {
                var typeElem = doc.GetElement(el.GetTypeId());
                if (typeElem is ElementType et)
                {
                    Parameter tp;
                    if (field.ParameterId != ElementId.InvalidElementId)
                        tp = et.get_Parameter(parBIC);
                    else
                        tp = et.LookupParameter(field.GetName());

                    if (tp != null)
                        p = tp;
                }
            }

            // Ошибка
            if (p == null)
                throw new Exception("Ошибка поиска параметра. Отправь разработчику!");


            if (!p.HasValue)
                return string.Empty;

            if (p.StorageType == StorageType.String)
                return p.AsString() ?? string.Empty;

            string vs = p.AsValueString();
            if (!string.IsNullOrEmpty(vs))
                return DoubleWithoutCulture(vs);

            return p.AsString() ?? string.Empty;
        }

        /// <summary>
        /// Стварае ключ групы для элемента (па sort/group палях).
        /// </summary>
        public static string BuildGroupKeyForElement(Document doc, Element el, IList<GroupFieldInfo> groupFields)
        {
            var parts = new List<string>(groupFields.Count);
            foreach (var gf in groupFields)
            {
                string v = GetFieldValueForElement(doc, el, gf.Field) ?? string.Empty;
                parts.Add(v);
            }

            return string.Join("~", parts);
        }

        /// <summary>
        /// Стварае ключ групы для радка спецыфікацыі (па тэкстах у неабходных калонках).
        /// rowCells – твае тэксты з WPF/мадэлі (без калонкі Use).
        /// </summary>
        public static string BuildGroupKeyForRow(IList<string> rowCells, IList<GroupFieldInfo> groupFields)
        {
            var parts = new List<string>(groupFields.Count);
            foreach (var gf in groupFields)
            {
                int c = gf.ColumnIndex;
                string v = c >= 0 && c < rowCells.Count ? rowCells[c] : string.Empty;
                parts.Add(v ?? string.Empty);
            }

            return string.Join("~", parts);
        }

        /// <summary>
        /// Стварае слоўнік група → элементы для ўсеx элементаў спецыфікацыі.
        /// </summary>
        public static Dictionary<string, List<ElementId>> BuildGroups(Document doc, ViewSchedule vs, IList<GroupFieldInfo> groupFields)
        {
            var dict = new Dictionary<string, List<ElementId>>();

            var elems = new FilteredElementCollector(doc, vs.Id)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (Element el in elems)
            {
                string key = BuildGroupKeyForElement(doc, el, groupFields);

                if (!dict.TryGetValue(key, out var list))
                {
                    list = new List<ElementId>();
                    dict.Add(key, list);
                }

                list.Add(el.Id);
            }

            return dict;
        }

        private static string DoubleWithoutCulture(string input)
        {
            if (double.TryParse(input.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double value))
                return value.ToString();

            return input;
        }
    }
}

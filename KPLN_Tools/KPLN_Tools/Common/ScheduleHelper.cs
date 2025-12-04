using Autodesk.Revit.DB;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
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
        public static ScheduleFormM ReadSchedule(ViewSchedule viewSchedule)
        {
            if (viewSchedule == null)
                throw new ArgumentNullException(nameof(viewSchedule));

            TableData tableData = viewSchedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);

            int rows = body.NumberOfRows;
            int cols = body.NumberOfColumns;

            var model = new ScheduleFormM
            {
                CurrentSchedule = viewSchedule
            };

            // 1) Загалоўкі слупкоў
            // Часта загалоўкі ў першай радку (0), але можна асобна браць Header-секцыю.
            for (int col = 0; col < cols; col++)
            {
                string headerText = body.GetCellText(0, col);
                if (string.IsNullOrEmpty(headerText))
                    headerText = $"Col {col + 1}";

                model.ColumnHeaders.Add(new CellData(headerText, false));
            }

            // 2) Радкі табліцы
            // Прапускаем радок загалоўкаў (0), ідзём з 1 да rows-1
            for (int row = 0; row < rows; row++)
            {
                CellType cellType = body.GetCellType(row, 0);
                bool isElementRow = (cellType == CellType.Parameter || cellType == CellType.CombinedParameter);
                
                var rowData = new List<CellData>();
                for (int col = 0; col < cols; col++)
                {
                    string cellText = viewSchedule.GetCellText(SectionType.Body, row, col);
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
                    ColumnIndex = i
                });
            }

            return result;
        }

        /// <summary>
        /// Вяртае значэнне поля для элемента, як максімальна падобнае да таго, што ў спецыфікацыі.
        /// </summary>
        public static string GetFieldValueForElement(Element el, ScheduleField field)
        {
            Parameter p = null;

            if (field.ParameterId != ElementId.InvalidElementId)
                p = el.get_Parameter((BuiltInParameter)field.ParameterId.IntegerValue);

            // Если не нашли по id
            if (p == null)
                p = el.LookupParameter(field.GetName());

            if (p == null)
                return string.Empty;

            if (p.StorageType == StorageType.String)
                return p.AsString() ?? string.Empty;

            string vs = p.AsValueString();
            if (!string.IsNullOrEmpty(vs))
                return vs;

            return p.AsString() ?? string.Empty;
        }

        /// <summary>
        /// Стварае ключ групы для элемента (па sort/group палях).
        /// </summary>
        public static string BuildGroupKeyForElement(Element el, IList<GroupFieldInfo> groupFields)
        {
            var parts = new List<string>(groupFields.Count);
            foreach (var gf in groupFields)
            {
                string v = GetFieldValueForElement(el, gf.Field) ?? string.Empty;
                parts.Add(v);
            }
            return string.Join("|", parts);
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

            return string.Join("|", parts);
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
                string key = BuildGroupKeyForElement(el, groupFields);

                if (!dict.TryGetValue(key, out var list))
                {
                    list = new List<ElementId>();
                    dict.Add(key, list);
                }

                list.Add(el.Id);
            }

            return dict;
        }

        private static int FindColumnIndexByHeader(TableSectionData body, int headerRow, string heading)
        {
            int cols = body.NumberOfColumns;
            for (int c = 0; c < cols; c++)
            {
                string text = body.GetCellText(headerRow, c);
                if (string.Equals(text, heading, StringComparison.OrdinalIgnoreCase))
                    return c;
            }

            return -1;
        }
    }
}

using Autodesk.Revit.DB;
using KPLN_Tools.Common;
using System.Collections.Generic;

namespace KPLN_Tools.Forms.Models
{
    public sealed class CellData
    {
        public CellData(string data, bool isEditableCell)
        {
            CD_Data = data;
            CD_IsEditableCell = isEditableCell;
        }

        public CellData(string data, bool isEditableCell, bool isCellChanged) : this(data, isEditableCell)
        {
            CD_IsCellChanged = isCellChanged;
        }

        /// <summary>
        /// Данные ячейки
        /// </summary>
        public string CD_Data { get; }
        
        /// <summary>
        /// Ячейка редактируется?
        /// </summary>
        public bool CD_IsEditableCell { get; }

        /// <summary>
        /// Ячейка была исправлена?
        /// </summary>
        public bool CD_IsCellChanged { get; set; }
    }

    public sealed class ScheduleFormM
    {
        /// <summary>
        /// Ссылка на спецификацию ревит
        /// </summary>
        public ScheduleEntity SE_Schedule { get; set; }

        /// <summary>
        /// Заголовки столбцов (0..ColumnCount-1)
        /// </summary>
        public List<CellData> ColumnHeaders { get; } = new List<CellData>();

        /// <summary>
        /// Строки таблицы. Каждая строка - отдельный лист
        /// </summary>
        public List<List<CellData>> Rows { get; } = new List<List<CellData>>();

        public int RowCount => Rows.Count;

        public int ColumnCount => ColumnHeaders.Count;
    }
}

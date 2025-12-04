using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Windows;
using UIFramework;

namespace KPLN_Tools.ExecutableCommand
{
    internal class ExcCmd_SchIncSetter : IExecutableCommand
    {
        private readonly ScheduleFormVM _scheduleForm;

        public ExcCmd_SchIncSetter(ScheduleFormVM scheduleForm)
        {
            _scheduleForm = scheduleForm;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            Document doc = uiDoc.Document;

            ScheduleFormM model = _scheduleForm.SFModel;
            ViewSchedule vs = model.SE_Schedule.SE_ViewSchedule;

            // 1. Атрымаем sort/group палі і калонкі
            List<GroupFieldInfo> groupFields = ScheduleHelper.GetGroupFields(vs);
            if (groupFields.Count == 0)
            {
                MessageBox.Show(_scheduleForm.MainSchIncWindow, "Не удалось собрать информацию по группированию/сортировке. Обратись к разработчику!", "Ошибка автоинкременты", MessageBoxButton.OK, MessageBoxImage.Error);
                return Result.Cancelled;
            }

            // 2. Пабудуем групы: ключ групы → элементы
            Dictionary<string, List<ElementId>> groups = ScheduleHelper.BuildGroups(doc, vs, groupFields);

            // 3. Даведаемся, у які параметр трэба пісаць з нашай калонкі
            TableData td = vs.GetTableData();
            TableSectionData body = td.GetSectionData(SectionType.Body);

            using (Transaction t = new Transaction(doc, "KPLN: Нумерация спецификаций"))
            {
                t.Start();

                // Забираю ID параметра для внесения измов и индекс столбца который отредактировал юзер
                ElementId targetParamId = null;
                int targetCol = -1;
                int rows = body.NumberOfRows;
                int cols = body.NumberOfColumns;
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        if (!model.Rows[row][col].CD_IsCellChanged)
                            continue;

                        targetParamId = body.GetCellParamId(row, col);
                        targetCol = col;
                        break;
                    }

                    if (targetParamId != null && targetCol != -1)
                        break;
                }


                // Записываю значения, предварительно верстая группирование элементов
                for (int modelRow = 0; modelRow < model.Rows.Count; modelRow++)
                {
                    List<CellData> rowCells = model.Rows[modelRow];

                    CellData targetCell = rowCells[targetCol];
                    if (!targetCell.CD_IsCellChanged)
                        continue;

                    // будуем ключ групы па тэкстах радка
                    var cellTexts = new List<string>(rowCells.Count);
                    foreach (CellData cd in rowCells)
                        cellTexts.Add(cd.CD_Data);

                    string key = ScheduleHelper.BuildGroupKeyForRow(cellTexts, groupFields);

                    if (!groups.TryGetValue(key, out var elemsInGroup) ||
                        elemsInGroup.Count == 0)
                        continue;

                    // запісваем у параметр каждого элемента групы
                    foreach (ElementId id in elemsInGroup)
                    {
                        Element el = doc.GetElement(id);
                        if (el == null)
                            continue;

                        Parameter p = null;
                        if (targetParamId != ElementId.InvalidElementId)
                            p = el.get_Parameter((BuiltInParameter)targetParamId.IntegerValue);

                        if (p == null)
                            p = el.LookupParameter(body.GetCellText(0, targetCol));

                        if (p == null || p.IsReadOnly)
                            continue;

                        p.Set(targetCell.CD_Data ?? string.Empty);
                    }
                }

                t.Commit();
            }

            // Закрываю все скрытые столбцы в спеке
            if (model.SE_Schedule.SE_HiddenFieldIds.Count > 0)
            {
                ScheduleDefinition def = model.SE_Schedule.SE_ViewSchedule.Definition;

                using (Transaction t = new Transaction(doc, "KPLN: Спецификация показать скрытые"))
                {
                    t.Start();

                    foreach (ScheduleFieldId fieldId in model.SE_Schedule.SE_HiddenFieldIds)
                    {
                        if (!def.IsValidFieldId(fieldId))
                            continue;

                        ScheduleField field = def.GetField(fieldId);
                        field.IsHidden = true;
                    }

                    model.SE_Schedule.SE_ViewSchedule.RefreshData();
                    t.Commit();
                }
            }

            return Result.Succeeded;
        }
    }
}

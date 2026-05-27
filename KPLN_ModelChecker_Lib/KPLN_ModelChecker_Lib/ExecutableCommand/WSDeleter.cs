using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace KPLN_ModelChecker_Lib.ExecutableCommand
{
    internal class WSDeleter : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly Dictionary<Workset, Workset> _wsReplacementMap;

        public WSDeleter(Document doc, Dictionary<Workset, Workset> wsReplacementMap)
        {
            _doc = doc;
            _wsReplacementMap = new Dictionary<Workset, Workset>();
            foreach (var kv in wsReplacementMap)
            {
                if (kv.Key.IsValidObject)
                    _wsReplacementMap[kv.Key] = kv.Value;
            }
        }

        public Result Execute(UIApplication app)
        {
#if !Debug2020 && !Revit2020
            // Проверим, не удаляем ли мы все пользовательские WS
            WorksetTable wsTable = _doc.GetWorksetTable();

            Workset[] allUserWs = Util.GetDocWorksets(_doc);
            if (_wsReplacementMap.Count >= allUserWs.Count())
            {
                MessageBox.Show("Нельзя удалить все пользовательские рабочие наборы.", 
                    "Ошибка",
                    MessageBoxButton.OK, 
                    MessageBoxImage.Warning);
                
                return Result.Cancelled;
            }


            // Если среди удаляемых — активный WS, переключимся на любой оставшийся
            WorksetId activeId = wsTable.GetActiveWorksetId();
            bool activeIsSelected = _wsReplacementMap.Keys.Any(w => w.Id == activeId);

            WorksetId fallbackActiveId = null;
            if (activeIsSelected)
            {
                var leftAfterDelete = allUserWs
                    .Where(ws => !_wsReplacementMap.Keys.Any(d => d.Id == ws.Id))
                    .Select(ws => ws.Id)
                    .ToList();

                if (!leftAfterDelete.Any())
                {
                    MessageBox.Show("Не удалось подобрать активный рабочий набор для переключения.", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return Result.Cancelled;
                }

                fallbackActiveId = leftAfterDelete.First();
            }


            // Переключение, если вдруг РН активный
            if (fallbackActiveId != null)
            {
                using (var tSwitch = new Transaction(_doc, "KPLN: Переключение активного РН"))
                {
                    tSwitch.Start();

                    wsTable.SetActiveWorksetId(fallbackActiveId);
                    
                    tSwitch.Commit();
                }

            }


            // Перевод в Редактируемый, чтобы можно было удалить
            ICollection<WorksetId> requestStateWSId = WorksharingUtils.CheckoutWorksets(_doc, _wsReplacementMap.Keys.Select(ws => ws.Id).ToArray());


            // Удаляю и информирую пользователя
            using (var tDel = new Transaction(_doc, $"KPLN: Удаление РН"))
            {
                tDel.Start();

                var errList = new List<string>();
                foreach (var kvp in _wsReplacementMap)
                {
                    Workset ws = kvp.Key;

                    // Пустой РН - переношу эл-ты в активный РН (удаление - опасно, хоть элементов и нет в РН)
                    DeleteWorksetSettings dws;
                    if (kvp.Value == null)
                        dws = new DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, wsTable.GetActiveWorksetId());
                    // Есть РН на замену
                    else
                        dws = new DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, kvp.Value.Id);
                    

                    if (WorksetTable.CanDeleteWorkset(_doc, ws.Id, dws))
                        WorksetTable.DeleteWorkset(_doc, ws.Id, dws);
                    else
                        errList.Add($"«{ws.Name}»");
                }


                // Вывожу ошибку пользователю
                if (errList.Any())
                {
                    MessageBox.Show($"Процесс прерван. Ошибка при удалении след. рабочих наборов:\n{string.Join(", ", errList)}.\n\n Обратись к разработчику",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    tDel.RollBack();
                    return Result.Cancelled;
                }
                else
                    MessageBox.Show($"Выбранные рабочие наборы - успешно удалены!" +
                            $"\n\nСписок рабочих наборов: {string.Join(", ", _wsReplacementMap.Select(kvp => kvp.Key.Name))}",
                        "Результат",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                tDel.Commit();
            }
#endif

            return Result.Succeeded;
        }
    }
}
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ModelChecker_Lib.ExecutableCommand
{
    internal class WSDeleter : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly IEnumerable<Workset> _wsCollection;

        public WSDeleter(Document doc, IEnumerable<Workset> wsColl)
        {
            _doc = doc;
            _wsCollection = wsColl.Where(el => el.IsValidObject);
        }

        public Result Execute(UIApplication app)
        {
            // Проверим, не удаляем ли мы все пользовательские WS
            WorksetTable wsTable = _doc.GetWorksetTable();

            Workset[] allUserWs = new FilteredWorksetCollector(_doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .ToArray();
            if (_wsCollection.Count() >= allUserWs.Count())
            {
                MessageBox.Show("Нельзя удалить все пользовательские рабочие наборы.", "Empty WS",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return Result.Cancelled;
            }

            // Если среди удаляемых — активный WS, переключимся на любой оставшийся
            WorksetId activeId = wsTable.GetActiveWorksetId();
            bool activeIsSelected = _wsCollection.Any(w => w.Id == activeId);

            WorksetId fallbackActive = null;
            if (activeIsSelected)
            {
                var leftAfterDelete = allUserWs
                    .Where(ws => !_wsCollection.Any(d => d.Id == ws.Id))
                    .Select(ws => ws.Id)
                    .ToList();

                if (!leftAfterDelete.Any())
                {
                    MessageBox.Show("Не удалось подобрать активный рабочий набор для переключения.", "Empty WS",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return Result.Cancelled;
                }

                fallbackActive = leftAfterDelete.First();
            }

            using (var tg = new TransactionGroup(_doc, "Удаление пустых рабочих наборов"))
            {
                tg.Start();

                try
                {
                    if (fallbackActive != null)
                    {
                        using (var tSwitch = new Transaction(_doc, "Переключение активного WS"))
                        {
                            tSwitch.Start();
                            wsTable.SetActiveWorksetId(fallbackActive);
                            tSwitch.Commit();
                        }
                    }

                    using (var t = new Transaction(_doc, $"KPLN: Удаление РН"))
                    {
                        foreach (Workset ws in _wsCollection)
                        {
                            t.Start();
                            try
                            {
                                ICollection<ElementId> delElems = _doc.Delete(new ElementId(ws.Id.IntegerValue));
                                if (!delElems.Any())
                                {
                                    // Revit вернул false — удаление не прошло
                                    MessageBox.Show($"Не удалось удалить «{ws.Name}».", "Empty WS",
                                        MessageBoxButton.OK, MessageBoxImage.Warning);
                                }
                                t.Commit();
                            }
                            catch (Exception ex)
                            {
                                t.RollBack();
                                MessageBox.Show($"Ошибка при удалении «{ws.Name}»:\n{ex.Message}", "Empty WS",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }

                    tg.Assimilate();
                }
                catch
                {
                    tg.RollBack();
                    throw;
                }
            }

            return Result.Succeeded;
        }
    }
}

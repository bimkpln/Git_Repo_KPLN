using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.ExecutableCommand
{
    public sealed class DeleteElemsExcCmd : IExecutableCommand
    {
        private readonly DeleteUnusedViewsM _entity;

        public DeleteElemsExcCmd(DeleteUnusedViewsM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;


            // Получаю список на удаление
            List<ElementId> elIds = new List<ElementId>();
            foreach (var tee in _entity.TreeElemEntities)
            {
                GetCheckedTEERElems(tee, elIds);
            }


            // Удаляю
            using (Transaction trans = new Transaction(doc, $"KPLN: {ExtCmdDeleteUnusedViews.PluginName}"))
            {
                trans.Start();

                int cnt = 0;
                List<string> errors = new List<string>();
                foreach (var id in elIds)
                {
                    try
                    {
                        doc.Delete(id);
                        cnt++;
                    }
                    catch (Exception ex)
                    {
                        if (app.ActiveUIDocument.ActiveView.Id == id)
                            errors.Add($"Элемент с id \"{id}\" имя \"{doc.GetElement(id).Name}\" - ошибка: " +
                                $"Это последний активный вид, его нельзя удалить. Перейди на другой вид и удали повторно");
                        else
                            errors.Add($"Элемент с id \"{id}\" имя \"{doc.GetElement(id).Name}\" - ошибка: {ex.Message}");
                        
                    }
                }

                if (errors.Any())
                {
                    MessageBox.Show(
                        _entity.MainWindow,
                        $"Из модели сейчас успешно удалится {cnt} шт. видов. " +
                            $"При удалении возникли ошибки, см. детально в отдельном окне (появиться после удаления)",
                        $"KPLN: {ExtCmdDeleteUnusedViews.PluginName}",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);

                    foreach (var error in errors)
                    {
                        HtmlOutput.Print(error, MessageType.Error);
                    }
                }
                else
                {
                    MessageBox.Show(
                        _entity.MainWindow,
                        $"Из модели сейчас успешно удалится {cnt} шт. видов",
                        $"KPLN: {ExtCmdDeleteUnusedViews.PluginName}",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию ревит-элементов, которые отметили в дереве
        /// </summary>
        /// <param name="tee"></param>
        /// <param name="result"></param>
        private static void GetCheckedTEERElems(TreeElementEntity tee, List<ElementId> result)
        {
            if (tee == null)
                return;

            if (tee.TEE_Element != null && tee.IsChecked)
                result.Add(tee.TEE_Element.Id);

            if (tee.TEE_ChildrenColl == null || tee.TEE_ChildrenColl.Count == 0)
                return;

            foreach (TreeElementEntity child in tee.TEE_ChildrenColl)
            {
                GetCheckedTEERElems(child, result);
            }
        }
    }
}

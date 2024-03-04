using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ViewsAndLists_Ribbon.Common;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandListTBlockParamCopier : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;

            //Workin with elements
            List<ViewSheet> mixedSheetsList = new List<ViewSheet>();
            List<Element> falseElemList = new List<Element>();
            List<ElementId> selIds = sel.GetElementIds().ToList();
            foreach (ElementId selId in selIds)
            {
                Element elem = doc.GetElement(selId);
                int catId = elem.Category.Id.IntegerValue;
                if (catId.Equals((int)BuiltInCategory.OST_Sheets))
                {
                    ViewSheet curViewSheet = elem as ViewSheet;
                    mixedSheetsList.Add(curViewSheet);
                }
                else
                {
                    falseElemList.Add(elem);
                }
            }

            //Main part of code
            if (mixedSheetsList.Count == 0)
            {
                TaskDialog.Show("Ошибка", "В выборке нет ни одного листа, или вообще ничего не выбрано :(", TaskDialogCommonButtons.Ok);
                return Result.Cancelled;
            }
            else
            {
                try
                {
                    List<TBlockEntity> tBlocksEntities = CreateTBlocksEntity(doc, mixedSheetsList);

                    using (Transaction trans = new Transaction(doc, "KPLN: Параметры осн.надписей"))
                    {
                        trans.Start();

                        foreach (TBlockEntity tBlockEntity in tBlocksEntities)
                        {
                            SetParamFromListToTBlock(tBlockEntity, "КП_Ш_Номер листа", "Номер листа вручную");
                            SetParamFromListToTBlock(tBlockEntity, "Имя листа", "Системный_Имя листа");
                            SetParamFromListToTBlock(tBlockEntity, "Номер листа", "Системный_Номер листа_Стандартный");
                            SetParamFromListToTBlock(tBlockEntity, "Орг.КомплектЧертежей", "Системный_Орг.КомплектЧертежей");
                        }

                        trans.Commit();
                    }
                }
                catch (UserErrorException uex)
                {
                    string msg = string.Empty;
                    if (uex.ErrorElements == null)
                        msg = $"ОШИБКА:\n{uex.ErrorMessage}";
                    else
                        msg = $"ОШИБКА:\n{uex.ErrorMessage}\nДля элементов:\n{string.Join(", ", uex.ErrorElements.Select(e => e.Id.ToString()))}";

                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Нет возможности выполнить")
                    {
                        MainContent = msg,
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning
                    };
                    taskDialog.Show();

                    return Result.Failed;
                }
                catch (Exception ex)
                {
                    Print($"Прервано с критической системной ошибкой:\n {ex.Message}", KPLN_Loader.Preferences.MessageType.Error);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        private List<TBlockEntity> CreateTBlocksEntity(Document doc, IEnumerable<ViewSheet> sheetsList)
        {
            List<TBlockEntity> result = new List<TBlockEntity>();

            foreach (ViewSheet sheet in sheetsList)
            {
                IList<ElementId> dependentElemsColl = sheet.GetDependentElements(null);
                IEnumerable<Element> tBlocksOnView = dependentElemsColl
                    .Select(id => doc.GetElement(id))
                    .Where(el => el.Category != null && el.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks);
                result.Add(new TBlockEntity(sheet, tBlocksOnView));
            }

            return result;
        }

        private void SetParamFromListToTBlock(TBlockEntity tBlockEntity, string listParamName, string tBlockParamName)
        {
            ViewSheet viewSheet = tBlockEntity.CurrentViewSheet;
            Parameter vSheetParam = viewSheet.LookupParameter(listParamName);
            if (vSheetParam == null)
                throw new UserErrorException($"У листа {viewSheet.SheetNumber} - {viewSheet.Name} нет параметра {listParamName}");

            IEnumerable<Parameter> tBlockParams = tBlockEntity.CurrentTBlocks.Select(tbl => tbl.LookupParameter(tBlockParamName));
            if (tBlockParams == null || tBlockParams.Contains(null))
            {
                Print($"У основной надписи на листе {viewSheet.SheetNumber} - {viewSheet.Name} нет параметра {tBlockParamName}. Если это титул - пропусти, иначе - скинь в BIM-отдел",
                    KPLN_Loader.Preferences.MessageType.Warning);
                return;
            }

            foreach (Element elem in tBlockEntity.CurrentTBlocks)
            {
                string vSheetParamValue = vSheetParam.AsString();
                if (elem is FamilyInstance famInst)
                {
                    if (famInst.Category.Id.IntegerValue != (int)BuiltInCategory.OST_TitleBlocks)
                        throw new Exception("Скинь разработчику - в коллекцию TBlockEntity попали НЕ только OST_TitleBlocks");

                    Parameter currentTBlockParam = famInst.LookupParameter(tBlockParamName);
                    if (currentTBlockParam.IsReadOnly)
                        throw new UserErrorException($"Параметр {tBlockParamName} у основной надписи с id: {famInst.Id} только для чтения. Перезаписать не получиться!");

                    currentTBlockParam.Set(vSheetParamValue);
                }
                else
                    throw new Exception("Скинь разработчику - ошибка апкастинга в FamilyInstance");

            }
        }
    }
}

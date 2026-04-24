using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SetParamsByFrameExtCmd : IExternalCommand
    {
        internal const string PluginName = "Выбрать/заполнить рамкой";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            try
            {
                // Выделенные эл-ты
                IEnumerable<Element> selectedElemsToFind = SelectionSearchFilter.UserSelectedFilters(uidoc);
                if (selectedElemsToFind == null || !selectedElemsToFind.Any())
                    return Result.Cancelled;


                SetParamsByFrameForm mainForm = new SetParamsByFrameForm(uiapp, selectedElemsToFind);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);

#if !Debug2020 && !Revit2020
                mainForm.Show();
#else
                mainForm.ShowDialog();
#endif
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }
    }
}

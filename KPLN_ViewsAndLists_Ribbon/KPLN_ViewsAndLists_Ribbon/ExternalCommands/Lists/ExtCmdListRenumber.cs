using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_PluginActivityWorker;
using KPLN_ViewsAndLists_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists
{
    internal class NumericStringComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (double.TryParse(x, NumberStyles.Any, CultureInfo.InvariantCulture, out double xValue) &&
                double.TryParse(y, NumberStyles.Any, CultureInfo.InvariantCulture, out double yValue))
            {
                return xValue.CompareTo(yValue);
            }

            return string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmdListRenumber : IExternalCommand
    {
        internal const string PluginName = "Перенумеровать листы";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;

            List<ViewSheet> mixedSheetsList = new List<ViewSheet>();
            List<Element> falseElemList = new List<Element>();
            List<ElementId> selIds = sel.GetElementIds().ToList();
            foreach (ElementId selId in selIds)
            {
                Element elem = doc.GetElement(selId);
                if (elem is ViewSheet curViewSheet)
                    mixedSheetsList.Add(curViewSheet);
                else
                    falseElemList.Add(elem);
            }
            List<ViewSheet> sortedSheets = mixedSheetsList.OrderBy(s => s.SheetNumber, new NumericStringComparer()).ToList();

            if (mixedSheetsList.Count != 0)
            {
                ParameterSet titleBlockParams = mixedSheetsList[0].Parameters;
                ListRenumberForm inputForm = new ListRenumberForm(uiapp, sortedSheets, titleBlockParams);
                if ((bool)inputForm.ShowDialog())
                {
                    if (falseElemList.Count != 0)
                    {
                        int cnt = 0;
                        foreach (Element elem in falseElemList)
                        {
                            cnt++;
                        }
                        string msg = string.Format("Если что, были случайно выбраны элементы, которые не являются листами. Успешно проигнорировано {0} штук/-и", cnt.ToString());
                        TaskDialog.Show("Предупреждение", msg, TaskDialogCommonButtons.Ok);
                    }

                    DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                    return Result.Succeeded;
                }

                return Result.Cancelled;
            }

            TaskDialog.Show("Ошибка", "В выборке нет ни одного листа, или вообще ничего не выбрано :(", TaskDialogCommonButtons.Ok);
            return Result.Cancelled;
        }
    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Views.Colorize;
using KPLN_ViewsAndLists_Ribbon.Views.Colorize.FilterData;
using KPLN_ViewsAndLists_Ribbon.Views.FilterUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CommandViewColoring : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                View curView = doc.ActiveView;

                //проверю, можно ли менять фильтры у данного вида
                bool checkAllowFilters = ViewUtils.CheckIsChangeFiltersAvailable(doc, curView);
                if (!checkAllowFilters)
                {
                    TaskDialog.Show("Ошибка", "Невозможно назначить фильтры, так как они определяются в шаблоне вида.");
                    return Result.Failed;
                }

                //список все параметров у элементов на виде
                List<Element> elems = new FilteredElementCollector(doc, curView.Id)
                     .WhereElementIsNotElementType()
                     .ToElements()
                     .Where(e => e != null)
                     .Where(e => e.IsValidObject)
                     .Where(e => e.Category != null)
                     .Where(e => e.Category.Id.IntegerValue != -2000500)
                     .ToList();

                List<MyParameter> mparams = ViewUtils.GetAllFilterableParameters(doc, elems);


                FormSelectParameterForFilters form1 = new FormSelectParameterForFilters
                {
                    parameters = mparams
                };
                if (form1.ShowDialog() != System.Windows.Forms.DialogResult.OK) return Result.Cancelled;


                IFilterData filterData = null;
                if (form1.colorizeMode == ColorizeMode.ByParameter)
                {
                    int startSymbols = 0;
                    if (form1.criteriaType == CriteriaType.StartsWith)
                    {
                        startSymbols = form1.startSymbols;
                    }
                    filterData = new FilterDataSimple(doc, elems, form1.selectedParameter, form1.startSymbols, form1.criteriaType);
                }
                else if (form1.colorizeMode == ColorizeMode.CheckHostmark)
                {
                    filterData = new FilterDataForRebars(doc);
                }

                MyDialogResult collectResult = filterData.CollectValues(doc, curView);
                if (collectResult.ResultType == ResultType.cancel)
                {
                    return Result.Cancelled;
                }
                else if (collectResult.ResultType == ResultType.error)
                {
                    message = collectResult.Message;
                    return Result.Failed;
                }
                else if (collectResult.ResultType == ResultType.warning)
                {
                    TaskDialog.Show("Внимание", collectResult.Message);
                }


                if (filterData.ValuesCount > 64)
                {
                    message = "Значений больше 64! Генерация цветов невозможна";
                    return Result.Failed;
                }

                //Получу id сплошной заливки
                ElementId solidFillPatternId = DocumentGetter.GetSolidFillPatternId(doc);


                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Колоризация вида");

                    foreach (ElementId filterId in curView.GetFilters())
                    {
                        ParameterFilterElement filter = doc.GetElement(filterId) as ParameterFilterElement;
                        if (filter.Name.StartsWith(filterData.FilterNamePrefix))
                        {
                            curView.RemoveFilter(filterId);
                        }
                    }

                    filterData.ApplyFilters(doc, curView, solidFillPatternId, form1.colorLines, form1.colorFill);


                    t.Commit();
                }

                return Result.Succeeded;

            }

            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");
                return Result.Failed;
            }

        }
    }
}

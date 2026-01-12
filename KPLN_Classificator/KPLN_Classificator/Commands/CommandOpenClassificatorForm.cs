using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Classificator.Data;
using KPLN_Classificator.Forms;
using KPLN_Classificator.Utils;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Classificator
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandOpenClassificatorForm : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ApplicationConfig.IsDocumentAvailable = true;
            StorageUtils utils = new StorageUtils(commandData.Application.Application.VersionNumber);
            ClassificatorForm form;

            if (commandData.Application.ActiveUIDocument == null)
            {
                ApplicationConfig.IsDocumentAvailable = false;

                form = new ClassificatorForm(utils, new List<MyParameter>() { });
                form.Show();

                return Result.Succeeded;
            }

            Document doc = commandData.Application.ActiveUIDocument.Document;
            LastRunInfo.createInstance(doc);

            List<BuiltInCategory> filteredCategorys = new List<BuiltInCategory>()
            {
                BuiltInCategory.OST_ProjectInformation
            };

            List<Element> elems = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements()
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                .Where(e => e != null && e.IsValidObject && e.Category != null && !filteredCategorys.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
#else
                .Where(e => e != null && e.IsValidObject && e.Category != null && !filteredCategorys.Contains(e.Category.BuiltInCategory))
#endif
                .ToList();

            List<MyParameter> mparams = ViewUtils.GetAllFilterableParameters(doc, elems);

            form = new ClassificatorForm(utils, mparams.Distinct().ToList());
            form.Show();

            return Result.Succeeded;
        }
    }
}

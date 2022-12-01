using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Classificator.Data;
using KPLN_Classificator.Forms;
using KPLN_Classificator.Utils;
using System;
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
            ApplicationConfig.isDocumentAvailable = true;
            StorageUtils utils = new StorageUtils();
            ClassificatorForm form;

            if (commandData.Application.ActiveUIDocument == null)
            {
                ApplicationConfig.isDocumentAvailable = false;
                
                form = new ClassificatorForm(utils, new List<MyParameter>() { });
                form.Show();
                
                return Result.Succeeded;
            }

            Document doc = commandData.Application.ActiveUIDocument.Document;
            LastRunInfo.createInstance(doc);

            List<int> filteredCategorys = new List<int>() 
            {
                //-2000500, 
                //-2003100, 
                //-2000280, 
                -2003101
            };

            List<Element> elems = new FilteredElementCollector(doc)
             .WhereElementIsNotElementType()
             .ToElements()
             .Where(e => e != null)
             .Where(e => e.IsValidObject)
             .Where(e => e.Category != null)
             .Where(e => !filteredCategorys.Contains(e.Category.Id.IntegerValue))
             .ToList();

            List<MyParameter> mparams = ViewUtils.GetAllFilterableParameters(doc, elems);
            
            form = new ClassificatorForm(utils, mparams.Distinct().ToList());
            form.Show();
            
            return Result.Succeeded;
        }
    }
}

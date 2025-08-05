using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandRemoveInstance : IExecutableCommand
    {
        public CommandRemoveInstance() {}

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;
            
            try
            {
                Document doc = uiDoc.Document;
                
                if (doc != null)
                {
                    Transaction t = new Transaction(doc, "KPLN_Указатель очистить");
                    t.Start();
                    
                    // Чистка от старых экз.
                    FamilyInstance[] oldFamInsOfGM = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_GenericModel)
                        .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == CommandPlaceFamily.FamilyName)
                        .Cast<FamilyInstance>()
                        .ToArray();
                        
                    ICollection<ElementId> availableWSOldElemsId = WorksharingUtils.CheckoutElements(doc, oldFamInsOfGM.Select(el => el.Id).ToArray());
                        
                    doc.Delete(availableWSOldElemsId);

                    t.Commit();
                }
            }
            catch (Exception)
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}

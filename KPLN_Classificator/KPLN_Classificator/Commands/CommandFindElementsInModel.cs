using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Classificator.Forms.ViewModels;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Classificator
{
    class CommandFindElementsInModel : IExecutableCommand
    {
        private List<RuleItem> ruleItems;

        public CommandFindElementsInModel(List<RuleItem> ruleItems)
        {
            this.ruleItems = ruleItems;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            List<Element> constrs = new List<Element>();
            foreach (RuleItem ruleItem in ruleItems)
            {
                List<Element> elements = getElemsFromModel(doc, ruleItem);
                constrs.AddRange(elements);
                ruleItem.elemsCountInModel = elements.Count;
            }
            Selection sel = app.ActiveUIDocument.Selection;
            sel.SetElementIds(constrs.Select(x => x.Id).ToList());
            return Result.Succeeded;
        }

        public static List<Element> getElemsFromModel(Document doc, RuleItem ruleItem)
        {
            List<BuiltInCategory> constrCats = new List<BuiltInCategory>() { ruleItem.builtInCategoryName };
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WherePasses(new ElementMulticategoryFilter(constrCats))
                .Where(x => ParamUtils.nameChecker(ruleItem.familyNameValue, ParamUtils.getElemFamilyName(x)))
                .Where(x => ParamUtils.nameChecker(ruleItem.typeNameValue, x.Name))
                .Where(x => ParamUtils.paramChecker(ruleItem.parameterName, ruleItem.parameterValue, x))
                .ToList();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Xml;
using Autodesk.Revit.UI.Selection;
using KPLN_Classificator.Data;
using KPLN_Classificator.Forms;
using Autodesk.Revit.DB.Architecture;
using System.Windows.Forms;
using KPLN_Classificator.Forms.ViewModels;
using KPLN_Loader.Common;

namespace KPLN_Classificator
{
    public class CommandGetElementInfo : IExecutableCommand
    {
        private MainWindow form;
        public List<string> paramNames;

        public CommandGetElementInfo(MainWindow form)
        {
            this.form = form;
            this.paramNames = form.settings.paramNameItems.Select(x => x.paramName).ToList();
        }
        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Selection sel = app.ActiveUIDocument.Selection;

            if (sel.GetElementIds().Count == 0)
            {
                MessageBox.Show("Для генерации правил на основе элемента, выберите хотя бы 1 элемент в модели", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return Result.Cancelled;
            }

            List<ElementId> elemsIds = sel.GetElementIds().ToList();
            int counter = 0;
            foreach (ElementId elemId in elemsIds)
            {
                Element elem = doc.GetElement(elemId);

                if (elem is Group) continue;

                BuiltInCategory builtInCategory;
                string familyName = null;
                string typeName = null;
                List<string> paramValues = new List<string>();

                if (elem is Room)
                {
                    familyName = elem.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString();
                }
                else
                {
                    if (famNameParam != null)
                        familyName = famNameParam.AsValueString();
                    Parameter famNameParam = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM);
                    familyName = familyName == null || familyName.Length == 0 ? (elem as ElementType).FamilyName : familyName;
                }

                typeName = elem.Name;

                builtInCategory = (BuiltInCategory)(elem.Category).Id.IntegerValue;

                foreach (string paramName in paramNames)
                {
                    if (paramName.Equals("") || paramName.Equals("Введите имя параметра"))
                    {
                        paramValues.Add("");
                    }
                    else
                    {
                        Parameter elemParam = elem.LookupParameter(paramName) ?? elem.Document.GetElement(elem.GetTypeId()).LookupParameter(paramName);
                        if (elemParam != null)
                            paramValues.Add(ParamUtils.GetValueStringOfParam_ByTargetParam(elem, paramName, elemParam) ?? "");
                    }
                }

                if (form.setRuleFromElement(builtInCategory, familyName, typeName, paramValues))
                {
                    form.ruleItems.Last().elemsCountInModel = CommandFindElementsInModel.getElemsFromModel(doc, form.ruleItems.Last()).Count;
                    counter++;
                }
            }
            form.findDoubledRules();
            MessageBox.Show(string.Format("Добавлено правил: {0}", counter), "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return Result.Succeeded;
        }
    }
}

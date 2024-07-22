using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.Forms;
using System;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OV_OZKDuctAccessory : IExternalCommand
    {
        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            //Get application and documnet objects
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FamilySymbol[] famSybols = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .WhereElementIsElementType()
                .Where(el => string.Compare(el.LookupParameter("Ключевая пометка").AsString(), "KPLN: Универсальный клапан ОЗК", StringComparison.OrdinalIgnoreCase) == 0)
                .Cast<FamilySymbol>()
                .ToArray();

            if (famSybols.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"В проекте отсутсвуют спец. семейства универсальных клапанов ОЗК (ищу по параметру \"Ключевая пометка\")",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);

                return Result.Cancelled;
            }
            else
            {
                OZKDuctAccessoryEntity[] ozkEntities = famSybols
                    .Select(fs => new OZKDuctAccessoryEntity(
                        fs,
                        fs
                        .GetDependentElements(null)
                        .Where(id => doc.GetElement(id) is FamilyInstance famInst && famInst.SuperComponent == null)
                        .Select(id => (FamilyInstance)doc.GetElement(id))
                        .ToArray()))
                    .ToArray();

                OV_OZKDuctAccessoryForm mainForm = new OV_OZKDuctAccessoryForm(uiapp, ozkEntities);
                mainForm.Show();
            }

            return Result.Succeeded;
        }
    }
}

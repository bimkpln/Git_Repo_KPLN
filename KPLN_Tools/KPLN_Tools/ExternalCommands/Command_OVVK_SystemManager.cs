using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OVVK_SystemManager : IExternalCommand
    {
        /// <summary>
        /// Коллекция BuiltInCategory используемых в моделях ОВВК типа FamilyInstance
        /// </summary>
        public static BuiltInCategory[] FamileInstanceBICs = new BuiltInCategory[]
        {
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_GenericModel,
        };

        /// <summary>
        /// Коллекция BuiltInCategory используемых в моделях ОВВК типа MEPCurve
        /// </summary>
        public static BuiltInCategory[] MEPCurveBICs = new BuiltInCategory[]
        {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_PipeInsulations,
        };

        /// <summary>
        /// Функция для фильтрации семейств в коллекциях FilteredElementCollector
        /// </summary>
        public static Func<FamilyInstance, bool> FamilyNameFilter = x =>
            !x.Symbol.FamilyName.StartsWith("500_")
            && !x.Symbol.FamilyName.StartsWith("501_")
            && !x.Symbol.FamilyName.StartsWith("502_")
            && !x.Symbol.FamilyName.StartsWith("503_");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Element> docElemsColl = new List<Element>();
            // Коллекция для анализа
            foreach (BuiltInCategory bic in Command_OVVK_SystemManager.FamileInstanceBICs)
            {
                docElemsColl.AddRange(new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(Command_OVVK_SystemManager.FamilyNameFilter));
            }

            foreach (BuiltInCategory bic in Command_OVVK_SystemManager.MEPCurveBICs)
            {
                docElemsColl.AddRange(new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType());
            }

            if (!docElemsColl.Any())
            {
                MessageBox.Show("В модели отсутсвуют элементы ОВВК для анализа!", "KPLN: Внимание", MessageBoxButton.OK);
                return Result.Cancelled;
            }

            OVVK_SystemManagerForm form = new OVVK_SystemManagerForm(docElemsColl);
            form.ShowDialog();

            return Result.Succeeded;
        }
    }
}

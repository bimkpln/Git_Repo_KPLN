using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OV_DuctThickness : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            // Коллекция для анализа
            IEnumerable<Element> ducts = new FilteredElementCollector(doc).OfClass(typeof(Duct)).WhereElementIsNotElementType();
            IEnumerable<Element> ductFittings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType();
            Element[] docElemsColl = ducts.Union(ductFittings).ToArray();

            if (docElemsColl.Length == 0)
            {
                MessageBox.Show("В модели отсутсвуют воздуховоды и соед. детали для анализа!", "KPLN: Внимание", MessageBoxButton.OK);
                return Result.Cancelled;
            }

            OV_DuctThicknessForm form = new OV_DuctThicknessForm(doc, docElemsColl);
            form.ShowDialog();

            return Result.Succeeded;
        }
    }
}

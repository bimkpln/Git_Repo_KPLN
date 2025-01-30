using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OV_DuctThickness : IExternalCommand
    {
        /// <summary>
        ///  GUID параметра для исключения перезаписи ("ТС_Перезаписать")
        /// </summary>
        internal static readonly Guid RevalueParamGuid = new Guid("466e6ecb-f390-43da-9cb5-76858d500a2c");
        
        internal const string PluginName = "ОВ: Толщина воздуховодов";

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

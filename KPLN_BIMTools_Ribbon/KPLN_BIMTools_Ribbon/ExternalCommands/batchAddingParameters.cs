using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Forms;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class batchAddingParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            string activeFamilyName;

            //Проверка на тип документа
            if (doc.IsFamilyDocument)
            {
                FamilyManager familyManager = doc.FamilyManager;
                activeFamilyName = doc.Title;
            }
            else
            {
                activeFamilyName = "Семейство не выбрано";
            }
         
            var window = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);        
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;  
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
            
            return Result.Succeeded;
        }
        
    }
}
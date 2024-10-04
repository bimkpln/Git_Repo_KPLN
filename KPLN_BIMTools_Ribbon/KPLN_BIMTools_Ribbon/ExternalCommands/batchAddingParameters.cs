using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_BIMTools_Ribbon.Forms;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class batchAddingParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            //Проверка на тип документа
            if (doc.IsFamilyDocument)
            {
                FamilyManager familyManager = doc.FamilyManager;
                string activeFamilyName = doc.Title;

                var window = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);
                var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                window.ShowDialog();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Текущий документ не является семейством.", "Предупреждение", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return Result.Failed;
            }


            return Result.Succeeded;
        }
        
    }
}
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using System;
using System.Linq;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SearchByIdExtCmd : IExternalCommand
    {
        internal const string PluginName = "ID-поиск в связях";
        private const string _mainViewNamePart = "KPLN_IDSearch";
        private SearchByIdForm _mainForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;

            // Создаю вид и открываю его
            View3D special3DView = CreateSpecialView(uiapp);

            // Создаю форму
            _mainForm = new SearchByIdForm(uiapp, special3DView);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(_mainForm);

            _mainForm.Show();

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            return Result.Succeeded;
        }

        private static View3D CreateSpecialView(UIApplication uiapp)
        {
            Document doc = uiapp.ActiveUIDocument.Document;
            string viewFullName = $"{_mainViewNamePart}_{KPLN_Loader.Application.CurrentRevitUser.SystemName}";

            View3D specialView = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => v.Name.Equals(viewFullName));

            
            // Уже есть - возвращаю вид
            if (specialView != null)
                return specialView;

            
            // Ещё нет - создаю вид
            try
            {
                Transaction t = new Transaction(doc, "KPLN_Создать 3D-вид");

                t.Start();

                ViewFamilyType vft3d = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                if (vft3d == null)
                    return null;

                View3D view3d = View3D.CreateIsometric(doc, vft3d.Id);
                view3d.Name = $"{_mainViewNamePart}_{KPLN_Loader.Application.CurrentRevitUser.SystemName}";

                t.Commit();

                return view3d;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return null;
            }
        }
    }
}

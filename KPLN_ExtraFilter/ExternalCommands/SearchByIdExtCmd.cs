using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Entities.SearchById;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class SearchByIdExtCmd : IExternalCommand
    {
        internal const string PluginName = "ID-поиск в связях";
        private const string _mainViewNamePart = "KPLN_IDSearch";
        private const double _initialSectionBoxHalfSize = 1.0;
        private SearchByIdForm _mainForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;

            // Создаю вид и открываю его
            View3D special3DView = CreateSpecialView(uiapp);
            if (special3DView == null)
                return Result.Cancelled;

            if (TryGetSelectedLinkedElement(uiapp, out SearchByIdEntity selectedLinkedElement))
            {
                uiapp.ActiveUIDocument.ActiveView = special3DView;
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectByIdExсCmd(selectedLinkedElement));

                // Счетчик факта запуска
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                return Result.Succeeded;
            }

            if (HasAnySelection(uiapp))
            {
                MessageBox.Show(
                    "Выбери элемент внутри связи или сними выделение для ручного поиска по ID.",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return Result.Cancelled;
            }

            // Создаю форму
            _mainForm = new SearchByIdForm(uiapp, special3DView);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(_mainForm);

            _mainForm.Show();

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            return Result.Succeeded;
        }

        private static bool TryGetSelectedLinkedElement(UIApplication uiapp, out SearchByIdEntity selectedLinkedElement)
        {
            selectedLinkedElement = null;

#if Debug2020 || Revit2020
            return false;
#else
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null)
                return false;

            Document doc = uidoc.Document;
            IList<Reference> selectedReferences = uidoc.Selection.GetReferences();
            foreach (Reference selectedReference in selectedReferences)
            {
                if (selectedReference.LinkedElementId == null || selectedReference.LinkedElementId == ElementId.InvalidElementId)
                    continue;

                RevitLinkInstance rli = doc.GetElement(selectedReference.ElementId) as RevitLinkInstance;
                Document linkDoc = rli?.GetLinkDocument();
                Element linkedElement = linkDoc?.GetElement(selectedReference.LinkedElementId);
                if (linkedElement == null)
                    continue;

                SearchByIdDocEntity linkDocEntity = new SearchByIdDocEntity(linkDoc, new Element[] { linkedElement }, rli);
                selectedLinkedElement = new SearchByIdEntity(linkDocEntity, linkedElement);
                return true;
            }

            return false;
#endif
        }

        private static bool HasAnySelection(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null)
                return false;

#if Debug2020 || Revit2020
            return uidoc.Selection.GetElementIds().Any();
#else
            return uidoc.Selection.GetReferences().Any() || uidoc.Selection.GetElementIds().Any();
#endif
        }

        private static void ApplyInitialSectionBox(View3D view3d)
        {
            BoundingBoxXYZ initialBox = new BoundingBoxXYZ()
            {
                Min = new XYZ(-_initialSectionBoxHalfSize, -_initialSectionBoxHalfSize, -_initialSectionBoxHalfSize),
                Max = new XYZ(_initialSectionBoxHalfSize, _initialSectionBoxHalfSize, _initialSectionBoxHalfSize)
            };

            view3d.IsSectionBoxActive = true;
            view3d.SetSectionBox(initialBox);
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
                ApplyInitialSectionBox(view3d);

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

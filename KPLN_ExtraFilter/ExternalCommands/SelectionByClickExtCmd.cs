using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Forms;
using KPLN_Library_Forms.Services;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectionByClickExtCmd : IExternalCommand
    {
        internal const string PluginName = "Выбрать по элементу";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Пользовательский элемент
            Element userSelElem = null;
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 1)
                userSelElem = doc.GetElement(selectedIds.FirstOrDefault());
#if Debug2020 || Revit2020
            else
                throw new System.Exception(
                    "Отправь разработчику: Ошибка предварительной проверки. " +
                    "Попало несколько элементов в выборку, хотя должен быть 1.");
#endif

            // Окно пользовательского ввода
            SelectionByClickForm mainForm = new SelectionByClickForm(uiapp, userSelElem);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);

            mainForm.Show();

            return Result.Succeeded;
        }
    }
}

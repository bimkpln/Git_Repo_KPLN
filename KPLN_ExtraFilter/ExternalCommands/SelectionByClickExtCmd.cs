using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectionByClickExtCmd : IExternalCommand
    {
        internal const string PluginName = "Выбрать по элементу";
        private SelectionByClickForm _mainForm;

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

            // Создание ExternalEvent для выделения эл-в
            SelectionChangedHandler selHandler = new SelectionChangedHandler();
            ExternalEvent selExtEv = ExternalEvent.Create(selHandler);

            // Создание ExternalEvent для отписки от SelectionChanged (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubscribeHandler unsubHandler = new UnsubscribeHandler() { Handler = OnSelectionChanged };
            ExternalEvent unsubExtEv = ExternalEvent.Create(unsubHandler);

            // Окно пользовательского ввода
            _mainForm = new SelectionByClickForm(doc);
            _mainForm.SetExternalEvent(selExtEv, selHandler);
            // Предустановка
            if(userSelElem != null)
            {
                _mainForm.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelDoc = doc;
                _mainForm.CurrentSelectionByClickVM.CurrentSelectionByClickM.UserSelElem = userSelElem;
            }

            // Подписываюсь на SelectionChanged
            uiapp.SelectionChanged += OnSelectionChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubExtEv.Raise();

            _mainForm.Show();

            return Result.Succeeded;
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => _mainForm.RaiseUpdate();
    }
}

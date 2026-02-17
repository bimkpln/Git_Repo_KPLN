using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_DefaultPanelExtension_Modify.ExternalEventHandler;
using KPLN_DefaultPanelExtension_Modify.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Loader.Common;
using System.Linq;

namespace KPLN_DefaultPanelExtension_Modify.ExecutableCommands
{
    internal class ExcCmdListVPPositionStart : IExecutableCommand
    {
        internal const string PluginName = "Положение вида";
        internal const string HelpUrl = "http://moodle/mod/book/view.php?id=502&chapterid=1344";

        private readonly Element[] _selElems;
        private ListVPPositionMainFrom _mainForm;

        public ExcCmdListVPPositionStart(Element[] selElems)
        {
            _selElems = selElems;
        }

        public Result Execute(UIApplication app)
        {
#if Debug2020 || Revit2020
            return Result.Cancelled;
        }
#else

            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;

            Element[] selTrueElems = _selElems
                .Where(el => el is Viewport || el is ScheduleSheetInstance)
                .ToArray();


            // Создание ExternalEvent для выделения эл-в
            ListVPPositionHandler handler = new ListVPPositionHandler();
            ExternalEvent extEv = ExternalEvent.Create(handler);


            // Создание ExternalEvent для отписки от выбора эл-в (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
            UnsubEventHandler unsubHandler = new UnsubEventHandler() { Handler = OnSelectionChanged };
            ExternalEvent unsubExtEv = ExternalEvent.Create(unsubHandler);


            // Окно пользовательского ввода
            _mainForm = new ListVPPositionMainFrom(selTrueElems);
            _mainForm.SetExternalEvent(extEv, handler);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(_mainForm);


            // Подписываюсь на SelectionChanged
            app.SelectionChanged += OnSelectionChanged;
            // Подписываю окно на отписку (через ExternalEvent)
            _mainForm.Closed += (s, e) => unsubExtEv.Raise();
            _mainForm.Show();


            return Result.Succeeded;
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => _mainForm.RaiseSelChanged();
#endif
    }
}

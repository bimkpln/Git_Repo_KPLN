using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandShowElement : IExecutableCommand
    {
        private readonly IEnumerable<Element> _elementCollection;

        public CommandShowElement(IEnumerable<Element> elemColl)
        {
            _elementCollection = elemColl;
        }

        public CommandShowElement(Element element)
        {
            _elementCollection = new List<Element>(1) { element };
        }

        public Result Execute(UIApplication app)
        {
            app.DialogBoxShowing += DialogBoxShowingEvant;

            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Демонстрация"))
            {
                t.Start();

                app.ActiveUIDocument.ShowElements(_elementCollection.FirstOrDefault());
                app.ActiveUIDocument.Selection.SetElementIds(_elementCollection.Select(e => e.Id).ToList());

                t.Commit();
            }

            app.DialogBoxShowing -= DialogBoxShowingEvant;

            return Result.Succeeded;
        }

        private void DialogBoxShowingEvant(object sender, DialogBoxShowingEventArgs e)
        {
            if (e is TaskDialogShowingEventArgs td && td.Message.Equals("Невозможно подобрать подходящий вид."))
            {
                // Закрываю окно
                e.OverrideResult(1);

                // Анлиз размеров (только их!), размещенных на легенде
                if (_elementCollection.FirstOrDefault() is Dimension dim)
                {
                    UIApplication app = sender as UIApplication;
                    View appView = app.ActiveUIDocument.ActiveView;
                    View dimView = dim.View;

                    if (appView == null && dimView == null)
                        TaskDialog.Show("KPLN", $"У размера с ID: {dim.Id} нет вида. Обратись в BIM-отдел!");

                    if (appView.Id != dimView.Id)
                    {
                        if (dim.View is ViewPlan viewPlan)
                        {
                            ReferenceArray refArray = dim.References;
                            StringBuilder stringBuilder = new StringBuilder(refArray.Size);
                            foreach (Reference refItem in refArray)
                            {
                                stringBuilder.Append($"{refItem.ElementId}/");
                            }

                            TaskDialog.Show("KPLN", $"Размер скрыт из-за скрытия элементов, на которые он размещен. " +
                                $"Чтобы его найти, нужно чтобы основы размера были видны на плане: {viewPlan.Name}." +
                                $"\nId элементов основы: {stringBuilder.ToString().TrimEnd('/')}");
                        }
                        else if (dim.View is View view && view.ViewType == ViewType.Legend)
                            TaskDialog.Show("KPLN", $"Открой легенду ({dimView.Name}) вручную.");
                    }
                }
            }
        }
    }
}

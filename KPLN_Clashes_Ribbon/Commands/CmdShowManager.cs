using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Clashes_Ribbon.ExternalEventHandler;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CmdShowManager : IExternalCommand
    {
        private ReportManagerForm _mainForm;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                DBProject dBProject = null;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;

                if (uidoc != null)
                {
                    Document doc = uidoc.Document;
                    string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
                    dBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, ModuleData.RevitVersion);
                }

                if (uidoc == null || dBProject == null)
                {
                    // Для пользователей бим-отдела - показываю все проекты, включая архивные
                    bool isBIMUser = DBMainService.CurrentUserDBSubDepartment.Id == 8;

                    ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm(null, ModuleData.RevitVersion, isBIMUser);
                    if ((bool)selectedProjectForm.ShowDialog())
                        dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
                    else
                        return Result.Cancelled;
                }


                // Создание ExternalEvent для переключения видов
                ViewActivatedHandler viewHandler = new ViewActivatedHandler();
                ExternalEvent viewExtEv = ExternalEvent.Create(viewHandler);

                // Создание ExternalEvent для отписки от переключения видов (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
                UnsubViewActHandler unsubViewHandler = new UnsubViewActHandler() { Handler = OnViewChanged };
                ExternalEvent unsubViewExtEv = ExternalEvent.Create(unsubViewHandler);


                // Настройки окна
                _mainForm = new ReportManagerForm(dBProject);
                _mainForm.SetExternalEvent(viewExtEv, viewHandler);
                // Искусственный запуск, чтобы засетить нужные данные
                _mainForm.RaiseUpdateViewChanged();


                // Подписываюсь на события
                commandData.Application.ViewActivated += OnViewChanged;
                
                // Подписываю окно на отписки (через ExternalEvent)
                _mainForm.Closed += (s, e) => unsubViewExtEv.Raise();

                // Запускаю окно
                _mainForm.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return Result.Cancelled;
            }
        }

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => _mainForm.RaiseUpdateViewChanged();
    }
}

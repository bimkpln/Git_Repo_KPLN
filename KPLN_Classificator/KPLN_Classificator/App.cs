using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Classificator.ExecutableCommand;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Classificator.ApplicationConfig;

namespace KPLN_Classificator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        public static string assemblyPath = "";

        public App()
        {
            //Конфигурирование приложения для работы в стандартной среде
            output = new StubOutput();
            commandEnvironment = new StandartCommandEnvironment();
        }

        public Result OnStartup(UIControlledApplication application)
        {
            assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            string tabName = "Addins";
            try { application.CreateRibbonTab(tabName); } catch { }

            string panelName = "Классификатор";
            RibbonPanel panel = null;
            List<RibbonPanel> tryPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName).ToList();
            if (tryPanels.Count == 0)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }
            else
            {
                panel = tryPanels.First();
            }

            PushButton btnHostMark = panel.AddItem(new PushButtonData(
                "ClassificatorCompleteCommand",
                "KPLN_Classificator",
                assemblyPath,
                "KPLN_Classificator.CommandOpenClassificatorForm")
                ) as PushButton;

            //register events
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);

            return Result.Succeeded;
        }

        public void OnIdling(object sender, IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIControlledApplication controlledApplication = sender as UIControlledApplication;
            StandartCommandEnvironment sce = commandEnvironment as StandartCommandEnvironment;
            while (sce.getQueue().Count != 0)
            {
                try
                {
                    sce.getQueue().Dequeue().Execute(uiapp);
                }
                catch (Exception e)
                {
                    output.PrintErr(e, "Ошибка в процессе выполнения внешней команды.");
                }
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}

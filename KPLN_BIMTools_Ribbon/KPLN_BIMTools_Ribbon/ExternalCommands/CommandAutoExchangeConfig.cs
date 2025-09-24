using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Forms;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;


namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandAutoExchangeConfig : IExternalCommand
    {
        internal const string PluginName = "Конф. автозапуска";

        public CommandAutoExchangeConfig()
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ElementEntity[] plugins = new ElementEntity[]
            {
                new ElementEntity(RevitDocExchangeEnum.Revit.ToString(), "Плагин \"RVT: Обмен\""),
                new ElementEntity(RevitDocExchangeEnum.Navisworks.ToString(), "Плагин \"NWC: Обмен\""),
            };
            ElementSinglePick elementSinglePick = new ElementSinglePick(plugins);
            if (!(bool)elementSinglePick.ShowDialog())
                return Result.Cancelled;
            RevitDocExchangeEnum selectedEnum = (RevitDocExchangeEnum)System.Enum.Parse(typeof(RevitDocExchangeEnum), elementSinglePick.SelectedElement.Name);


            ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm(ModuleData.RevitVersion, true);
            if (!(bool)selectedProjectForm.ShowDialog())
                return Result.Cancelled;


            DBProject dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
            ConfigDispatcher configDispatcher = new ConfigDispatcher(dBProject, selectedEnum, true);
            if (!(bool)configDispatcher.ShowDialog())
                return Result.Cancelled;

            return Result.Succeeded;
        }
    }
}

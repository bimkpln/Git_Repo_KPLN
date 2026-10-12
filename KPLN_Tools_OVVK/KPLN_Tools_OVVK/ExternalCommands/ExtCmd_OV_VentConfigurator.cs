using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Tools_OVVK.Forms;
using KPLN_Tools_OVVK.Services.VentConfigurator;
using System;

namespace KPLN_Tools_OVVK.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class ExtCmd_OV_VentConfigurator : IExternalCommand
    {
        public const string PluginName = VentConfiguratorSettings.PluginName;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (OV_VentConfiguratorForm.TryActivateExisting())
                return Result.Succeeded;

            OV_VentConfiguratorForm form = null;
            try
            {
                form = new OV_VentConfiguratorForm();
                WindowHandleSearch.MainWindowHandle.SetAsOwner(form);
                form.StartTracking(commandData.Application);
                form.Show();
                form.Activate();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                form?.ReleaseResources();
                message = ex.ToString();
                return Result.Failed;
            }
        }
    }
}

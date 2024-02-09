using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using static KPLN_Loader.Output.Output;
using static KPLN_Loader.Preferences;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandExtraMonitoring_SetParams : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly ObservableCollection<MonitorParamRule> _paramRule;
        private readonly Dictionary<ElementId, List<MonitorEntity>> _monitorEntities;

        public CommandExtraMonitoring_SetParams(
            Document doc,
            ObservableCollection<MonitorParamRule> paramRule,
            Dictionary<ElementId, List<MonitorEntity>> monitorEntities)
        {
            _doc = doc;
            _paramRule = paramRule;
            _monitorEntities = monitorEntities;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Перенос параметров"))
            {
                t.Start();

                foreach (var kvp in _monitorEntities)
                {
                    RevitLinkInstance revitLinkInstance = _doc.GetElement(kvp.Key) as RevitLinkInstance;
                    foreach (var monitorEntity in kvp.Value)
                    {
                        Element targetElement = monitorEntity.ModelElement;
                        Element sourceElement = monitorEntity.LinkElement;
                        foreach (var paramRule in _paramRule)
                        {
                            Parameter targetParam = targetElement.GetParameters(paramRule.SelectedTargetParameter.Definition.Name).FirstOrDefault();
                            Parameter sourceParam = sourceElement.GetParameters(paramRule.SelectedSourceParameter.Definition.Name).FirstOrDefault();
                            if (sourceParam != null && targetParam != null)
                            {
                                switch (targetParam.StorageType)
                                {
                                    case StorageType.Double:
                                        double? dv = MonitorParamDataTool.GetDoubleValue(sourceParam);
                                        if (dv != null)
                                            targetParam.Set((double)dv);
                                        break;
                                    case StorageType.Integer:
                                        int? iv = MonitorParamDataTool.GetIntegerValue(sourceParam);
                                        if (iv != null)
                                            targetParam.Set((int)iv);
                                        break;
                                    case StorageType.ElementId:
                                        string eiv = MonitorParamDataTool.GetStringValue(sourceParam);
                                        if (eiv != null && eiv != " " && eiv != string.Empty)
                                            targetParam.Set(eiv);
                                        break;
                                    case StorageType.String:
                                        string sv = MonitorParamDataTool.GetStringValue(sourceParam);
                                        if (sv != null && sv != " " && sv != string.Empty)
                                            targetParam.Set(sv);
                                        break;
                                }
                            }

                            else
                                Print($"Проверь наличие параметра {paramRule.SelectedTargetParameter.Definition.Name} у элемента модели ({targetElement.Id}), " +
                                    $"или параметра {paramRule.SelectedSourceParameter.Definition.Name} у элемента связи ({sourceElement.Id})", MessageType.Error);
                        }
                    }
                }

                t.Commit();
            }

            new TaskDialog("Выполнено")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainContent = "Перенос завершен!",
            }
                .Show();

            return Result.Succeeded;
        }
    }
}
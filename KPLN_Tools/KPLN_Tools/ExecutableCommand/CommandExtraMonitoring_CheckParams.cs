using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandExtraMonitoring_CheckParams : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly ObservableCollection<MonitorParamRule> _paramRule;
        private readonly Dictionary<ElementId, List<MonitorEntity>> _monitorEntities;

        public CommandExtraMonitoring_CheckParams(
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
            // Коллекция ошибок, при отработке
            List<string> _localErrors = new List<string>();

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
                            string targetParamData = GetStringDataFromParam(targetParam);
                            string sourceParamData = GetStringDataFromParam(sourceParam);
                            if (!targetParamData.Equals(sourceParamData))
                            {
                                string errorMsg = $"У элемента твоей модели парамтер '{targetParam.Definition.Name}' имеет значение '{targetParamData}', " +
                                    $"а параметр из связи '{sourceParam.Definition.Name}' имеет значение '{sourceParamData}'. " +
                                    $"Id твоего элемента: {targetElement.Id}, Id элемента из связи: {sourceElement.Id}";
                                _localErrors.Add(errorMsg);
                            }
                        }

                        else
                            Print($"Проверь наличие параметра {paramRule.SelectedTargetParameter.Definition.Name} у элемента модели ({targetElement.Id}), " +
                                $"или параметра {paramRule.SelectedSourceParameter.Definition.Name} у элемента связи ({sourceElement.Id})", MessageType.Error);
                    }
                }
            }

            if (_localErrors.Count > 0)
            {
                foreach (string msg in _localErrors)
                {
                    Print($"{msg}", MessageType.Error);
                }
                Print($"Список ошибок в данных (не эквивалентны значения):", MessageType.Error);
            }
            else
            {
                new TaskDialog("Выполнено")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Проверка завершена, ошибок нет!",
                }
                .Show();
            }

            return Result.Succeeded;
        }
        private string GetStringDataFromParam(Parameter param)
        {
            switch (param.StorageType)
            {
                case StorageType.ElementId:
                    return param.AsValueString();
                case StorageType.Double:
                    return param.AsValueString();
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.String:
                    return param.AsValueString();
            }

            return null;
        }
    }
}

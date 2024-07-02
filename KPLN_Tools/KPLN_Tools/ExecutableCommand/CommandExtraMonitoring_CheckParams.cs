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
                    Element modelElement = monitorEntity.ModelElement;
                    Element linkElement = monitorEntity.CurrentMonitorLinkEntity.LinkElement;
                    foreach (var paramRule in _paramRule)
                    {
                        Parameter modelParam = modelElement.GetParameters(paramRule.SelectedTargetParameter.Definition.Name).FirstOrDefault();
                        Parameter linkParam = linkElement.GetParameters(paramRule.SelectedSourceParameter.Definition.Name).FirstOrDefault();
                        if (linkParam != null && modelParam != null)
                        {
                            string modelParamData = GetStringDataFromParam(modelParam);
                            string linkParamData = GetStringDataFromParam(linkParam);
                            if (!modelParamData.Equals(linkParamData))
                            {
                                string errorMsg = $"У элемента твоей модели парамтер '{modelParam.Definition.Name}' имеет значение '{modelParamData}', " +
                                    $"а параметр из связи '{linkParam.Definition.Name}' имеет значение '{linkParamData}'. " +
                                    $"Id твоего элемента: {modelElement.Id}, Id элемента из связи: {linkElement.Id}";
                                _localErrors.Add(errorMsg);
                            }
                        }

                        else
                            Print($"Проверь наличие параметра {paramRule.SelectedTargetParameter.Definition.Name} у элемента модели ({modelElement.Id}), " +
                                $"или параметра {paramRule.SelectedSourceParameter.Definition.Name} у элемента связи ({linkElement.Id})", MessageType.Error);
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
                    return param.AsString();
            }

            return null;
        }
    }
}

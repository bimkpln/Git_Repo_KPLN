using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExternalCommands;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandSystemManager_MergeSystem : IExecutableCommand
    {
        private readonly OVVK_SystemManager_VM _viewModel;
        private readonly OVVK_MergeSystem[] _sysDataToMerge;

        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _errorDict = new Dictionary<string, List<ElementId>>();
        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _warningDict = new Dictionary<string, List<ElementId>>();

        public CommandSystemManager_MergeSystem(OVVK_SystemManager_VM viewModel, OVVK_MergeSystem[] sysDataToMerge)
        {
            _viewModel = viewModel;
            _sysDataToMerge = sysDataToMerge;
        }

        public Result Execute(UIApplication app)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{Command_OVVK_SystemManager.PluginName}_Объединение систем по параметру", ModuleData.ModuleName).ConfigureAwait(false);

            using (Transaction t = new Transaction(_viewModel.CurrentDoc, $"KPLN: Объединение систем"))
            {
                t.Start();

                foreach (Element elem in _viewModel.ElementColl)
                {
                    Parameter elemSysParam = elem.LookupParameter(_viewModel.ParameterName);
                    if (elemSysParam == null)
                    {
                        AddToErrorDict(_warningDict, $"У элемента нет параметра \"{_viewModel.ParameterName}\"", elem.Id);
                        continue;
                    }

                    if (elemSysParam.IsReadOnly)
                    {
                        if (elem is FamilyInstance fi && fi.SuperComponent == null)
                            AddToErrorDict(_warningDict, $"У элемента параметра \"{_viewModel.ParameterName}\" - ТОЛЬКО для чтения. Нужно исправить семейство", elem.Id);
                        
                        continue;
                    }

                    string elemSysParamData = elemSysParam.AsString();
                    if (string.IsNullOrEmpty(elemSysParamData))
                        continue;
                    
                    foreach (OVVK_MergeSystem dataToMerge in _sysDataToMerge)
                    {
                        string newSysName = $"{dataToMerge.FirstParamName}/{dataToMerge.LastParamName}";
                        if (elemSysParamData.Equals(dataToMerge.FirstParamName) || elemSysParamData.Equals(dataToMerge.LastParamName))
                            elemSysParam.Set(newSysName);
                    }
                }

                t.Commit();
            }

            if (_warningDict.Keys.Count != 0)
            {
                PrintErrors(_warningDict);
                MessageBox.Show(
                    $"Скрипт отработал, но есть замечания к некоторым элементам. Список выведен отдельным окном",
                    "Внимание",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show(
                    $"Скрипт отработал без ошибок",
                    "Успешно",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Добавить эл-т ревит к замечанию
        /// </summary>
        /// <param name="errMsg"></param>
        /// <param name="elementId"></param>
        private void AddToErrorDict(Dictionary<string, List<ElementId>> dict, string errMsg, ElementId elementId)
        {
            if (dict.ContainsKey(errMsg))
            {
                List<ElementId> ids = dict[errMsg];
                ids.Add(elementId);

                dict[errMsg] = ids;
            }
            else
                dict.Add(errMsg, new List<ElementId> { elementId });
        }

        /// <summary>
        /// Вывод ошибок пользователю
        /// </summary>
        private void PrintErrors(Dictionary<string, List<ElementId>> dict)
        {
            foreach (KeyValuePair<string, List<ElementId>> kvp in dict)
            {
                string errorElemsId = string.Join(", ", kvp.Value);
                HtmlOutput.Print($"ОШИБКА: \"{kvp.Key}\" у элементов: {errorElemsId}", MessageType.Error);
            }
        }
    }
}

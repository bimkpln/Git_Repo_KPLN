using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.ExternalCommands;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class ExcCmd_SystemManager_MergeSystem : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly OVVK_MergeSystem[] _sysDataToMerge;

        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _warningDict = new Dictionary<string, List<ElementId>>();

        public ExcCmd_SystemManager_MergeSystem(OVVK_SystemManagerM viewModel, OVVK_MergeSystem[] sysDataToMerge)
        {
            _model = viewModel;
            _sysDataToMerge = sysDataToMerge;
        }

        public Result Execute(UIApplication app)
        {
            // Одна исходная система может входить только в одну группу за запуск.
            var replacementNames = new Dictionary<string, string>(StringComparer.Ordinal);
            var availableNames = new HashSet<string>(_model.SystemSumParameters, StringComparer.Ordinal);
            foreach (OVVK_MergeSystem group in _sysDataToMerge)
            {
                string[] names = group?.SelectedSystemNames ?? new string[0];
                if (!OVVK_MergeSystem.IsAllowedCount(names.Length) || names.Distinct().Count() != names.Length
                    || names.Any(name => string.IsNullOrWhiteSpace(name) || !availableNames.Contains(name) || replacementNames.ContainsKey(name)))
                {
                    _model.SetOperationMessage("Объединение отменено: проверь группы. Системы должны существовать и не повторяться.");
                    return Result.Cancelled;
                }

                string newName = SystemNameResolver.Normalize(
                    string.Join(_model.SysNameSeparator, names), _model.SysNameSeparator);
                foreach (string name in names)
                    replacementNames.Add(name, newName);
            }
            if (replacementNames.Count == 0)
                return Result.Cancelled;


            var readOnlyElements = new List<Element>();
            using (Transaction t = new Transaction(_model.CurrentDoc, $"KPLN: Объединение систем"))
            {
                if (t.Start() != TransactionStatus.Started)
                    return Result.Cancelled;

                foreach (Element elem in _model.ElementColl)
                {
                    Parameter elemSysParam = elem.LookupParameter(_model.ParameterName);
                    if (elemSysParam == null || elemSysParam.StorageType != StorageType.String)
                        continue;

                    string elemSysParamData = elemSysParam.AsString();
                    if (string.IsNullOrEmpty(elemSysParamData)
                        || !replacementNames.TryGetValue(elemSysParamData, out string newSysName))
                        continue;

                    if (elemSysParam.IsReadOnly)
                    {
                        readOnlyElements.Add(elem);
                        continue;
                    }

                    if (!elemSysParam.Set(newSysName))
                        throw new InvalidOperationException("Не удалось изменить параметр «" + _model.ParameterName
                            + "» элемента " + elem.Id + ". Объединение отменено.");
                }

                if (t.Commit() != TransactionStatus.Committed)
                {
                    _model.SetOperationMessage("Объединение систем отменено: Revit не подтвердил изменения.");
                    return Result.Cancelled;
                }
            }

            // Вложенные элементы могут получить новое значение от родителя при регенерации.
            // Сообщаем только о тех параметрах, которые действительно остались без изменения.
            foreach (Element element in readOnlyElements)
            {
                string value = element.LookupParameter(_model.ParameterName)?.AsString();
                if (value != null && replacementNames.ContainsKey(value))
                    AddToErrorDict(_warningDict, $"Параметр \"{_model.ParameterName}\" доступен только для чтения; система элемента не объединена", element.Id);
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

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ExtraFilter.ExecutableCommand
{
    internal class SetParamsByFrameExcCmd : IExecutableCommand
    {
        private readonly Element[] _elemsToSet;
        private readonly MainItem[] _currentParamEntities;
        /// <summary>
        /// Коллекция элементов с проблемами в назначении систем
        /// </summary>
        private Dictionary<string, List<Element>> _warningsElementColl = new Dictionary<string, List<Element>>();

        public SetParamsByFrameExcCmd(SetParamsByFrameEntity formEntity)
        {
            _elemsToSet = formEntity.SelectedElems;
            _currentParamEntities = formEntity.MainItems.ToArray();
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Запись конфигурации последнего запуска
            ConfigService.SaveConfig<MainItem>(ModuleData.RevitVersion, doc, ConfigType.Memory, _currentParamEntities);

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(SetParamsByFrameExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            using (Transaction trans = new Transaction(doc, "KPLN: Параметризация"))
            {
                trans.Start();

                foreach (Element elem in _elemsToSet)
                {
                    foreach (MainItem mainItem in _currentParamEntities)
                    {
                        string paramName = mainItem.UserSelectedParamEntity.CurrentParamName;
                        string newValue = mainItem.UserInputParamValue;

                        Parameter currentParam = elem.LookupParameter(paramName)
                            ?? elem.Document.GetElement(elem.GetTypeId()).LookupParameter(paramName);
                        if (currentParam == null)
                        {
                            AddToWarnings($"Отсутствует параметр {paramName} (как в типе, так и в экземпляре)", elem);
                            continue;
                        }

                        switch (currentParam.StorageType)
                        {
                            case StorageType.Double:
                                if (double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double dValue))
                                {
#if Debug2020 || Revit2020
                                    DisplayUnitType unitType = currentParam.DisplayUnitType;
                                    double prjData = UnitUtils.ConvertToInternalUnits(dValue, unitType);
#else
                                    ForgeTypeId forgeTypeId = currentParam.GetUnitTypeId();
                                    double prjData = UnitUtils.ConvertToInternalUnits(dValue, forgeTypeId);
#endif
                                    currentParam.Set(prjData);
                                    break;
                                }
                                goto case StorageType.Integer;
                            case StorageType.Integer:
                                if (int.TryParse(newValue, out int iValue))
                                {
                                    currentParam.Set(iValue);
                                    break;
                                }
                                goto case StorageType.String;
                            case StorageType.String:
                                currentParam.Set(newValue);
                                break;
                            default:
                                AddToWarnings($"Для параметра {paramName} был выбран не верный тип ввода (вместо цифры - текст)", elem);
                                break;
                        }
                    }
                }

                trans.Commit();
            }

            // Оставляю экстровыбор пользователю
            app.ActiveUIDocument.Selection.SetElementIds(_elemsToSet.Select(e => e.Id).ToList());

            if (_warningsElementColl.Count > 0)
            {
                foreach (KeyValuePair<string, List<Element>> kvp in _warningsElementColl)
                {
                    HtmlOutput.Print($"{kvp.Key}. Id эл-в с ошибкой: {string.Join(", ", kvp.Value.Select(elem => elem.Id.IntegerValue))}", MessageType.Warning);
                }
            }

            return Result.Succeeded;
        }

        private void AddToWarnings(string msg, Element elem)
        {
            if (_warningsElementColl.ContainsKey(msg))
                _warningsElementColl[msg].Add(elem);
            else
                _warningsElementColl[msg] = new List<Element>() { elem };

        }
    }
}

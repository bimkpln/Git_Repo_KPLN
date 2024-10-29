using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools.Common.OVVK_System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandDuctThickness_Start : IExecutableCommand
    {
        private readonly DuctThicknessEntity _currentDuctThicknessEntity;
        private readonly Element[] _elementsToSet;

        private readonly ExtensibleStorageBuilder _extensibleStorageBuilder;

        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _errorDict = new Dictionary<string, List<ElementId>>();

        public CommandDuctThickness_Start(DuctThicknessEntity ductThicknessEntity, Element[] elementsToSet)
        {
            _currentDuctThicknessEntity = ductThicknessEntity;
            _elementsToSet = elementsToSet;

            _extensibleStorageBuilder = new ExtensibleStorageBuilder(
                new Guid("753380C4-DF00-40F8-9745-D53F328AC139"),
                "Last_Run",
                "KPLN_DuctSize");
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"KPLN: Толщина воздуховодов"))
            {
                t.Start();

                #region Работа с ExtensibleStorage
                ProjectInfo pi = app.ActiveUIDocument.Document.ProjectInformation;
                Element piElem = pi as Element;
                _extensibleStorageBuilder.SetStorageData_TimeRunLog(piElem, app.Application.Username, DateTime.Now);
                #endregion

                if (!SetDuctThicknessData())
                {
                    t.RollBack();
                    return Result.Cancelled;
                }
                else
                    t.Commit();
            }

            if (_errorDict.Keys.Count != 0)
            {
                PrintErrors();
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
        private void AddToErrorDict(string errMsg, ElementId elementId)
        {
            if (_errorDict.ContainsKey(errMsg))
            {
                List<ElementId> ids = _errorDict[errMsg];
                ids.Add(elementId);

                _errorDict[errMsg] = ids;
            }
            else
                _errorDict.Add(errMsg, new List<ElementId> { elementId });
        }

        /// <summary>
        /// Вывод ошибок пользователю
        /// </summary>
        private void PrintErrors()
        {
            foreach (KeyValuePair<string, List<ElementId>> kvp in _errorDict)
            {
                string errorElemsId = string.Join(", ", kvp.Value);
                HtmlOutput.Print($"ОШИБКА: \"{kvp.Key}\" у элементов: {errorElemsId}", MessageType.Error);
            }
        }

        /// <summary>
        /// Запись параметров в указанные элементы
        /// </summary>
        private bool SetDuctThicknessData()
        {
            foreach (Element elem in _elementsToSet)
            {
                Parameter paramToSet = elem.LookupParameter(_currentDuctThicknessEntity.ParameterName);
                if (paramToSet != null)
                {
                    ConnectorManager cManager = null;
                    if (elem is Duct ductElem)
                        cManager = ductElem.ConnectorManager;
                    else if (elem is FamilyInstance finstElem)
                        cManager = finstElem.MEPModel.ConnectorManager;
                    else
                    {
                        MessageBox.Show(
                            $"Не удалось привести тип к Duct или FamilyInstance у эл-та с id: {elem.Id}",
                            "ОШИБКА",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return false;
                    }

                    Parameter insulutionType = elem.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE);
                    Parameter systemType = elem.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM);

                    if (systemType == null || systemType.AsString().Contains("Не определено") || systemType.AsString().Contains("Нет системы"))
                        AddToErrorDict("Не определена система. Толщина НЕ записана", elem.Id);
                    else
                    {
                        List<Connector> roundConn = new List<Connector>();
                        List<Connector> rectConn = new List<Connector>();

                        ConnectorSet connectors = cManager.Connectors;
                        foreach (Connector connector in connectors)
                        {
                            if (connector.Shape == ConnectorProfileType.Round)
                                roundConn.Add(connector);
                            else if (connector.Shape == ConnectorProfileType.Rectangular)
                                rectConn.Add(connector);
                            else
                            {
                                MessageBox.Show(
                                    $"Отправь в BIM-отдел: Неконтролируемая форма соединителя у эл-та с id: {elem.Id}",
                                    "ОШИБКА",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                                return false;
                            }
                        }

                        if (roundConn.Count > 0 && rectConn.Count > 0)
                        {
                            double maxRoundConnRadius = roundConn.Max(c => c.Radius);
                            double maxRoundConn = maxRoundConnRadius * 2;
                            double maxRectConn = rectConn.Select(c => Math.Max(c.Height, c.Width)).Max();
                            if (maxRoundConn > maxRectConn)
                                SetThicknessData_RoundConnector(paramToSet, insulutionType, systemType, maxRoundConn);
                            else
                                SetThicknessData_RectangularConnector(paramToSet, insulutionType, systemType, maxRectConn);
                        }
                        else if (roundConn.Count > 0)
                        {
                            double maxRoundConnRadius = roundConn.Max(c => c.Radius);
                            double maxRoundConn = maxRoundConnRadius * 2;
                            SetThicknessData_RoundConnector(paramToSet, insulutionType, systemType, maxRoundConn);
                        }
                        else
                        {
                            double maxRectConn = rectConn.Select(c => Math.Max(c.Height, c.Width)).Max();
                            SetThicknessData_RectangularConnector(paramToSet, insulutionType, systemType, maxRectConn);
                        }
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"Работа экстренно завершена, т.к. отсутсвует параметр {_currentDuctThicknessEntity.ParameterName} у элемента с id: {elem.Id}",
                        "ОШИБКА",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;
        }

        private void SetThicknessData_RoundConnector(Parameter paramToSet, Parameter insulutionType, Parameter systemType, double maxSize)
        {
            double maxSize_mm = maxSize * 304.8;
            // Воздуховоды/соед. детали в огнезащ. изоляции, или противодымных систем
            if ((insulutionType != null 
                    && !string.IsNullOrEmpty(insulutionType.AsString()) 
                    && insulutionType.AsString().ToLower().Contains(_currentDuctThicknessEntity.PartOfInsulationName.ToLower()))
                || systemType.AsString().ToLower().Contains(_currentDuctThicknessEntity.PartOfSystemName.ToLower()))
            {
                if (maxSize_mm < 900)
                    paramToSet.Set(0.8 / 304.8);
                else if (maxSize_mm >= 900 && maxSize_mm < 1400)
                    paramToSet.Set(1.0 / 304.8);
                else if (maxSize_mm >= 1400 && maxSize_mm < 1800)
                    paramToSet.Set(1.2 / 304.8);
                else
                    paramToSet.Set(1.4 / 304.8);
            }
            // Воздуховоды/соед. детали общеобменных систем 
            else
            {
                if (maxSize_mm < 250)
                    paramToSet.Set(0.5 / 304.8);
                else if (maxSize_mm >= 250 && maxSize_mm < 500)
                    paramToSet.Set(0.6 / 304.8);
                else if (maxSize_mm >= 500 && maxSize_mm < 900)
                    paramToSet.Set(0.7 / 304.8);
                else if (maxSize_mm >= 900 && maxSize_mm < 1400)
                    paramToSet.Set(1.0 / 304.8);
                else if (maxSize_mm >= 1400 && maxSize_mm < 1800)
                    paramToSet.Set(1.2 / 304.8);
                else
                    paramToSet.Set(1.4 / 304.8);
            }
        }

        private void SetThicknessData_RectangularConnector(Parameter paramToSet, Parameter insulutionType, Parameter systemType, double maxSize)
        {
            double maxSize_mm = maxSize * 304.8;
            // Воздуховоды / соед.детали в огнезащ. изоляции, или противодымных систем
            if ((insulutionType != null
                    && !string.IsNullOrEmpty(insulutionType.AsString())
                    && insulutionType.AsString().ToLower().Contains(_currentDuctThicknessEntity.PartOfInsulationName.ToLower()))
                || systemType.AsString().ToLower().Contains(_currentDuctThicknessEntity.PartOfSystemName.ToLower()))
            {
                if (maxSize_mm < 1250)
                    paramToSet.Set(0.8 / 304.8);
                else
                    paramToSet.Set(0.9 / 304.8);
            }
            // Воздуховоды/соед. детали общеобменных систем 
            else
            {
                if (maxSize_mm < 300)
                    paramToSet.Set(0.5 / 304.8);
                else if (maxSize_mm >= 300 && maxSize_mm < 1250)
                    paramToSet.Set(0.7 / 304.8);
                else
                    paramToSet.Set(0.9 / 304.8);
            }
        }
    }
}

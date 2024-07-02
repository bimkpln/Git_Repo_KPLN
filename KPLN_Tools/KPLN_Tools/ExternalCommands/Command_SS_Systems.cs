using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common.SS_System;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Shapes;
using static Autodesk.Revit.DB.SpecTypeId;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_SS_Systems : IExternalCommand
    {
        private static readonly string _sysParamName = "КП_И_Адрес текущий";
        private static readonly string _previousAdressParamName = "КП_И_Адрес предыдущий";
        private static readonly string _adressParamName = "КП_И_Количество занимаемых адресов";
        private static readonly string _positionParamName = "КП_О_Позиция";
        private static readonly string _sysPrefixParamName = "КП_И_Префикс системы";
        private static readonly string _currentAdressParamName = "КП_И_Адрес устройства";
        private static readonly string _lenghtParamName = "КП_И_Длина линии";

        private ExternalCommandData _commandData;
        private UIApplication _uiapp;
        private UIDocument _uidoc;
        private Document _doc;

        /// <summary>
        /// Коллекция используемых параметров
        /// </summary>
        private readonly string[] _paramNames = new string[]
        {
            _sysParamName,
            _positionParamName,
            _sysPrefixParamName,
            _adressParamName,
            _previousAdressParamName,
            _currentAdressParamName,
            _lenghtParamName,

        };

        /// <summary>
        /// Коллекция типов линий в проекте
        /// </summary>
        private readonly Dictionary<string, GraphicsStyle> _lineStyles = new Dictionary<string, GraphicsStyle>();

        /// <summary>
        /// Коллекция типов систем СС
        /// </summary>
        private Dictionary<string, ElectricalSystemType> _electricalSystemTypes = new Dictionary<string, ElectricalSystemType>
        {
            { "Пожарная сигнализация", ElectricalSystemType.FireAlarm },
            { "Силовая система", ElectricalSystemType.PowerCircuit },
            { "Данные", ElectricalSystemType.Data },
            { "Проводные устройства связи", ElectricalSystemType.Telephone },
            { "Охранная сигнализация", ElectricalSystemType.Security },
            { "Вызов медперсонала", ElectricalSystemType.NurseCall },
            { "Управление", ElectricalSystemType.Controls },
        };

        private SS_Sysetm_Main _mainWindow;
        private SS_SystemEntity _previosSystemEntity;
        private SS_SystemEntity _currentSystemEntity;

        /// <summary>
        /// Queue для команд на выполнение по событию "OnIdling"
        /// </summary>
        public static Queue<ICommand> OnIdling_ICommandQueue = new Queue<ICommand>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _commandData = commandData;

            _uiapp = _commandData.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            _uiapp.Idling += Uiapp_Idling;

            #region Подготовка и создание формы
            IEnumerable<GraphicsStyle> styleCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(GraphicsStyle))
                .Cast<GraphicsStyle>();
            foreach (GraphicsStyle style in styleCollector)
            {
                if (style.GraphicsStyleCategory != null
                    && style.GraphicsStyleCategory.Parent != null
                    && style.GraphicsStyleCategory.Parent.Id.IntegerValue == (int)BuiltInCategory.OST_Lines)
                {
                    if (!_lineStyles.ContainsKey(style.Name))
                        _lineStyles.Add(style.Name, style);
                }
            }

            SS_SystemViewEntity sS_SystemViewEntity = new SS_SystemViewEntity
            {
                LineStyles = _lineStyles,
                ElectricalSystemTypes = _electricalSystemTypes,
                SystemNumber = "1.1",
                StartNumber = 1,
                SelectedSystemType = ElectricalSystemType.FireAlarm,
            };

            _mainWindow = new SS_Sysetm_Main(
                sS_SystemViewEntity,
                new SS_SystemCommand(AddParamsAction),
                new SS_SystemCommand(CreateConsistSystemAction),
                new SS_SystemCommand(CreateParallelSystemAction));
            _mainWindow.Show();
            #endregion

            //uiapp.Application.DocumentChanged += (sender, args) =>
            //{
            //    uiapp.Idling -= Uiapp_Idling;
            //};

            return Result.Succeeded;
        }

        private void Uiapp_Idling(object sender, IdlingEventArgs e)
        {
            while (OnIdling_ICommandQueue.Count != 0)
                OnIdling_ICommandQueue.Dequeue().Execute(_commandData);
        }

        /// <summary>
        /// Метод добавления параметров в проект
        /// </summary>
        private void AddParamsAction()
        {
            Autodesk.Revit.ApplicationServices.Application app = _uiapp.Application;

            CategorySet catSet = app.Create.NewCategorySet();
            catSet.Insert(_doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalCircuit));
            catSet.Insert(_doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalEquipment));

            app.SharedParametersFilename = @"X:\BIM\4_ФОП\КП_Файл общих парамеров.txt";
            DefinitionFile fopFile = app.OpenSharedParameterFile();
            List<Definition> fopDefintionColl = new List<Definition>(_paramNames.Length);
            foreach (DefinitionGroup defGroup in fopFile.Groups)
            {
                Definitions defs = defGroup.Definitions;
                if (defGroup.Name == "01 Обязательные ОБЩИЕ")
                    fopDefintionColl.AddRange(defs.Where(d => _paramNames.Contains(d.Name)));
                else if (defGroup.Name == "04 Обязательные ИНЖЕНЕРИЯ")
                    fopDefintionColl.AddRange(defs.Where(d => _paramNames.Contains(d.Name)));
            }

            if (fopDefintionColl.Count() < _paramNames.Length)
                throw new Exception("Не удалось получить ВСЕ необходимые параметры из ФОП");

            BindingMap bindMap = _doc.ParameterBindings;
            DefinitionBindingMapIterator binMapIterator = bindMap.ForwardIterator();
            binMapIterator.Reset();

            string msg = string.Empty;
            using (Transaction trans = new Transaction(_doc, "KPLN: Добавление параметров"))
            {
                trans.Start();
                bool isInserted = false;
                foreach (Definition def in fopDefintionColl)
                {
                    // Тут если параметр уже есть у части категорий из CategorySet - будет false. Нужно проработать
                    var newInstBind = app.Create.NewInstanceBinding(catSet);
                    if (_doc.ParameterBindings.Insert(def, newInstBind, BuiltInParameterGroup.PG_ENERGY_ANALYSIS))
                        isInserted = true;
                }
                if (isInserted)
                    msg = "Параметры добавлены. Можно приступать к заполнению";
                else
                    msg = "Параметры были добавлены ранее. Можно приступать к заполнению";

                trans.Commit();
            }

            TaskDialog td = new TaskDialog("ВНИМАНИЕ")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainInstruction = msg,
                CommonButtons = TaskDialogCommonButtons.Ok,
            };
            td.Show();
        }

        /// <summary>
        /// Выбор элементов в проекте
        /// </summary>
        /// <param name="uiapp"></param>
        private SS_SystemEntity ElemPicker(UIApplication uiapp)
        {
            try
            {
                ElementId selectedId = _uidoc
                    .Selection
                    .PickObject(ObjectType.Element, new PickFilter(), "Выбери следующий элемент. По окончанию - нажми \"Esc\"!")
                    .ElementId;
                SS_SystemEntity resultSysEntity = new SS_SystemEntity(_doc.GetElement(selectedId));
                if (resultSysEntity.CurrentElSystemSet.Count() != 0)
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Внимание!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                        MainContent = "Элемент уже был подключен к системе, добавить его в текущую, удалив старую?",
                        CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                    };

                    TaskDialogResult tdResult = taskDialog.Show();
                    if (tdResult == TaskDialogResult.Ok)
                    {
                        foreach (ElectricalSystem sys in resultSysEntity.CurrentElSystemSet)
                        {
                            _doc.Delete(sys.Id);
                        }
                    }
                }

                return resultSysEntity;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        /// <summary>
        /// Проверка формы на заполненность парамтеров перед стартом
        /// </summary>
        /// <returns>True если ошибка есть</returns>
        private bool IsPrepareFormError()
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSS_SystemViewEntity;
            if (viewEntity.SelectedSystemType == ElectricalSystemType.UndefinedSystemType)
            {
                TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainContent = "Нужно выбрать тип системы",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                taskDialog.Show();

                return true;
            }

            if (viewEntity.IsLineDraw == true && viewEntity.SelectedStyle == null)
            {
                TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainContent = "Если нужно построение цепи, то нужно выбрать тип линии",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                taskDialog.Show();

                return true;
            }

            return false;
        }

        /// <summary>
        /// Заполнение праматеров "КП_И_Адрес текущий", "КП_И_Адрес предыдущий" для ЭЛЕМЕНТОВ
        /// </summary>
        /// <returns></returns>
        private bool SetElementSystemData()
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSS_SystemViewEntity;

            #region "КП_И_Адрес текущий"
            Parameter currentSysNameParam = _currentSystemEntity.CurrentFamInst.LookupParameter(_sysParamName) ??
                throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentElem.Id} нет параметра {_sysParamName}");
            
            Parameter currentElemPositionParam = _currentSystemEntity.CurrentFamInst.LookupParameter(_positionParamName) ??
                _currentSystemEntity.CurrentFamInst.Symbol.LookupParameter(_positionParamName) ??
                throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentElem.Id} нет параметра {_positionParamName}");

            currentSysNameParam
                .Set($"{currentElemPositionParam.AsString()}{viewEntity.UserSeparator}{viewEntity.UserSystemIndex}{viewEntity.UserSeparator}{viewEntity.StartNumber}");
            #endregion

            #region "КП_И_Адрес предыдущий"
            if (_previosSystemEntity != null)
            {
                Parameter previousAdressParam = _currentSystemEntity.CurrentFamInst.LookupParameter(_previousAdressParamName) ??
                    throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentElem.Id} нет параметра {_previousAdressParamName}");
            
                Parameter previousSysNameParam = _previosSystemEntity.CurrentFamInst.LookupParameter(_sysParamName) ??
                    throw new Exception($"У элемента с id: {_previosSystemEntity.CurrentElem.Id} нет параметра {_sysParamName}");

                previousAdressParam.Set(previousSysNameParam.AsString());
            }
            #endregion

            return false;
        }

        /// <summary>
        /// Заполнение праматеров "КП_И_Адрес текущий", "КП_И_Адрес предыдущий", "КП_И_Префикс системы" для ЦЕПЕЦ
        /// </summary>
        /// <returns></returns>
        private bool SetSystemSystemData(ElectricalSystem elSys)
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSS_SystemViewEntity;

            #region "КП_И_Адрес текущий"
            Parameter currentSysNameParam = elSys.LookupParameter(_sysParamName) ??
                throw new Exception($"У цепи с id: {elSys.Id} нет параметра {_sysParamName}");

            Parameter currentElemPositionParam = _currentSystemEntity.CurrentFamInst.LookupParameter(_positionParamName) ??
                _currentSystemEntity.CurrentFamInst.Symbol.LookupParameter(_positionParamName) ??
                throw new Exception($"У элемена с id: {elSys.Id} нет параметра {_positionParamName}");

            currentSysNameParam
                .Set($"{currentElemPositionParam.AsString()}{viewEntity.UserSeparator}{viewEntity.UserSystemIndex}{viewEntity.UserSeparator}{viewEntity.StartNumber}");
            #endregion

            #region "КП_И_Адрес предыдущий"
            Parameter previousAdressParam = elSys.LookupParameter(_previousAdressParamName) ??
               throw new Exception($"У цепи с id: {_currentSystemEntity.CurrentElem.Id} нет параметра {_previousAdressParamName}");

            Parameter previousElemAdressParam = _currentSystemEntity.CurrentFamInst.LookupParameter(_sysParamName) ??
                throw new Exception($"У элемена с id: {_previosSystemEntity.CurrentElem.Id} нет параметра {_sysParamName}");

            previousAdressParam.Set(previousElemAdressParam.AsString());
            #endregion

            #region "КП_И_Префикс системы"
            Parameter sysPrefixParam = elSys.LookupParameter(_sysPrefixParamName) ??
               throw new Exception($"У цепи с id: {_currentSystemEntity.CurrentElem.Id} нет параметра {_sysPrefixParamName}");

            sysPrefixParam.Set(_mainWindow.CurrentSS_SystemViewEntity.UserSystemIndex);
            #endregion


            return false;
        }

        /// <summary>
        /// Построение линий между элементами цепи
        /// </summary>
        /// <returns></returns>
        private bool DrawSystemLines()
        {
            if (_previosSystemEntity.CurrentFamInst.Location is LocationPoint prevLocPnt 
                && _currentSystemEntity.CurrentFamInst.Location is LocationPoint currLocPnt)
            {
                List<XYZ> resultPnts = new List<XYZ>();
                XYZ prevPnt = prevLocPnt.Point;
                XYZ currPnt = currLocPnt.Point;

                // Проверяем направление прямой и строим перпендикуляр, если нужно
                if (AreEqual(prevPnt.X, currPnt.X, 0.1))
                    resultPnts.Add(new XYZ(prevPnt.X, currPnt.Y, 0));
                else if (AreEqual(prevPnt.Y, currPnt.Y, 0.1))
                    resultPnts.Add(new XYZ(currPnt.X, prevPnt.Y, 0));
                else
                {
                    /// Определяем координаты вектора между двумя точками
                    double dx = currPnt.X - prevPnt.X;
                    double dy = currPnt.Y - prevPnt.Y;

                    // Рассчитываем новые координаты для перпендикулярной линии
                    double newX, newY;
                    if (Math.Abs(dx) <= Math.Abs(dy))
                    {
                        // Если разница между координатами X меньше разницы между координатами Y,
                        // то линия параллельна оси Y
                        newX = prevPnt.X;
                        newY = currPnt.Y;
                    }
                    else
                    {
                        // Иначе линия параллельна оси X
                        newX = currPnt.X;
                        newY = prevPnt.Y;
                    }

                    resultPnts.Add(new XYZ(prevPnt.X, prevPnt.Y, 0));
                    resultPnts.Add(new XYZ(newX, newY, 0));
                    resultPnts.Add(new XYZ(currPnt.X, currPnt.Y, 0));
                }

                if (resultPnts.Count == 1)
                {
                    var line = Autodesk.Revit.DB.Line.CreateBound(prevPnt, resultPnts[0]);
                    DetailCurve detCurve = _doc.Create.NewDetailCurve(_doc.ActiveView, line);
                    detCurve.LineStyle = _mainWindow.CurrentSS_SystemViewEntity.SelectedStyle;
                }
                else
                {
                    for (int i = 0; i < resultPnts.Count - 1; i++)
                    {
                        var line = Autodesk.Revit.DB.Line.CreateBound(resultPnts[i], resultPnts[i + 1]);
                        DetailCurve detCurve = _doc.Create.NewDetailCurve(_doc.ActiveView, line);
                        detCurve.LineStyle = _mainWindow.CurrentSS_SystemViewEntity.SelectedStyle;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Проверка координат на эквивалентность
        /// </summary>
        /// <param name="coord1">Координата 1</param>
        /// <param name="coord2">Координата 2</param>
        /// <param name="epsilon">Допуск на совпадение</param>
        private bool AreEqual(double coord1, double coord2, double epsilon)
        {
            return Math.Abs(coord1 - coord2) < epsilon;
        }

        /// <summary>
        /// Создание последовательной цепи
        /// </summary>
        private void CreateConsistSystemAction()
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSS_SystemViewEntity;
            if (IsPrepareFormError())
                return;

            while (true)
            {

                _currentSystemEntity = ElemPicker(_uiapp);
                if (_currentSystemEntity == null)
                {
                    _mainWindow.CurrentSS_SystemViewEntity.StartNumber = 1;
                    break;
                }

                try
                {
                    using (Transaction trans = new Transaction(_doc, "KPLN: Параметры эл-та"))
                    {
                        trans.Start();
                        
                        SetElementSystemData();
                        
                        trans.Commit();
                    }
                        
                    if (_previosSystemEntity != null)
                    {
                        using (Transaction trans = new Transaction(_doc, "KPLN: Создание и параметры цепи"))
                        {
                            trans.Start();
                            
                            ElectricalSystem newElSys = ElectricalSystem
                                .Create(_doc, new List<ElementId> { _currentSystemEntity.CurrentElem.Id }, viewEntity.SelectedSystemType);
                            newElSys.SelectPanel(_previosSystemEntity.CurrentFamInst);
                            SetSystemSystemData(newElSys);

                            if (_mainWindow.CurrentSS_SystemViewEntity.IsLineDraw)
                                DrawSystemLines();
                            
                            trans.Commit();
                        }
                    }

                    _previosSystemEntity = _currentSystemEntity;
                        
                    _mainWindow.CurrentSS_SystemViewEntity.StartNumber++;
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainContent = "Элемент не содержит необходимого соединителя. Нужно исправить семейство",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    taskDialog.Show();
                }
                catch (Exception ex)
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainContent = $"Устрани ошибку: {ex.Message}",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    taskDialog.Show();
                }
            }


            _currentSystemEntity = null;
            _previosSystemEntity = null;
        }

        /// <summary>
        /// Создание параллельной
        /// </summary>
        private void CreateParallelSystemAction()
        {
            UIApplication uiapp = _commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


        }
    }
}

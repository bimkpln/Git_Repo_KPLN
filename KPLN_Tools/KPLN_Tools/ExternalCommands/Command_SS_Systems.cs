using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common.SS_System;
using KPLN_Tools.Forms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_SS_Systems : IExternalCommand
    {
        /// <summary>
        /// Queue для команд на выполнение по событию "OnIdling"
        /// </summary>
        public static Queue<ICommand> OnIdling_ICommandQueue = new Queue<ICommand>();

        private ExternalCommandData _commandData;
        private UIApplication _uiapp;
        private UIDocument _uidoc;
        private Document _doc;

        /// <summary>
        /// Коллекция типов линий в проекте
        /// </summary>
        private readonly Dictionary<string, GraphicsStyleWrapper> _lineStyles = new Dictionary<string, GraphicsStyleWrapper>();

        /// <summary>
        /// Коллекция типов систем СС
        /// </summary>
        private readonly Dictionary<string, ElectricalSystemType> _electricalSystemTypes = new Dictionary<string, ElectricalSystemType>
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
                        _lineStyles.Add(style.Name, new GraphicsStyleWrapper (style));
                }
            }

            SS_SystemViewEntity sS_SystemViewEntity = new SS_SystemViewEntity(_lineStyles, _electricalSystemTypes);

            _mainWindow = new SS_Sysetm_Main(
                sS_SystemViewEntity,
                new SS_SystemCommand(AddParamsAction),
                new SS_SystemCommand(CreateConsistSystemAction),
                new SS_SystemCommand(AddToConsistSystemPanelAction),
                new SS_SystemCommand(CreateParallelSystemAction),
                new SS_SystemCommand(RefreshSystemAction));

            _mainWindow.Show();
            #endregion

            return Result.Succeeded;
        }

        private void Uiapp_Idling(object sender, IdlingEventArgs e)
        {
            while (OnIdling_ICommandQueue.Count != 0)
                OnIdling_ICommandQueue.Dequeue().Execute(_commandData);
        }

        #region Основные Action для кнопок
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
            List<Definition> fopDefintionColl = new List<Definition>(SS_SystemEntity.EntityParams.Length);
            foreach (DefinitionGroup defGroup in fopFile.Groups)
            {
                Definitions defs = defGroup.Definitions;
                if (defGroup.Name == "01 Обязательные ОБЩИЕ")
                    fopDefintionColl.AddRange(defs.Where(d => SS_SystemEntity.EntityParams.Contains(d.Name)));
                else if (defGroup.Name == "04 Обязательные ИНЖЕНЕРИЯ")
                    fopDefintionColl.AddRange(defs.Where(d => SS_SystemEntity.EntityParams.Contains(d.Name)));
            }

            if (fopDefintionColl.Count() < SS_SystemEntity.EntityParams.Length)
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
        /// Создание последовательной цепи
        /// </summary>
        private void CreateConsistSystemAction()
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSystemViewEntity;
            if (IsPrepareFormError())
                return;

            while (true)
            {
                string msgToPick;
                if (_previosSystemEntity == null)
                    msgToPick = "Выбери первый элемент. Чтобы прервать - нажми \"Esc\"!";
                else
                    msgToPick = "Выбери следующий элемент. По закончить цикл выбора - нажми \"Esc\"!";
                _currentSystemEntity = ElemPicker(true, msgToPick);
                
                if (_currentSystemEntity == null) break;

                try
                {
                    using (Transaction trans = new Transaction(_doc, "KPLN_CS: Цепь"))
                    {
                        trans.Start();

                        SetElementSystemData();

                        if (_previosSystemEntity != null)
                        {

                            #region Создаю цепь
                            ConnectorSet mepConnectors = _currentSystemEntity.CurrentFamInst.MEPModel.ConnectorManager.Connectors;
                            Connector currentSysTypeConn = null;
                            foreach (Connector conn in mepConnectors)
                            {
                                if (conn.Domain == Domain.DomainElectrical
                                    && conn.ElectricalSystemType == viewEntity.SelectedSystemType 
                                    && conn.AllRefs.Size == 0)
                                {
                                    currentSysTypeConn = conn;
                                    break;
                                }
                            }

                            if (currentSysTypeConn == null)
                                throw new Exception($"У семейства ID: {_currentSystemEntity.CurrentFamInst.Id} - не осталось свободного соединителя нужноного типа");

                            ElectricalSystem elSys = ElectricalSystem.Create(currentSysTypeConn, viewEntity.SelectedSystemType);
                            #endregion

                            elSys.SelectPanel(_previosSystemEntity.CurrentFamInst);

                            SetSystemSystemData(elSys);

                            if (_mainWindow.CurrentSystemViewEntity.IsLineDraw)
                                DrawSystemLines();
                        }

                        _previosSystemEntity = _currentSystemEntity;
                        _mainWindow.CurrentSystemViewEntity.StartNumber++;
                        
                        trans.Commit();
                    }
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainContent = "Варианты ошибок:" +
                        "\n1. Элемент не содержит необходимого соединителя (нужно исправить семейство);" +
                        "\n2. Вы пытаетесь выполнить подключение с игнорированием последовательности (одна за одной) создания цепей;" +
                        "\n3. Вы пытаетесь подключить одну последовательную цепь в другую, что невозможно - у последовательной цепи ОДНА начальная панель." +
                        "\n\n Отмените 3 предыдущих действия (ctrl+z)!",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    taskDialog.Show();
                }
                catch (Exception ex)
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainContent = $"Устрани ошибку, или отправь разработчику: {ex.Message}",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    taskDialog.Show();
                }
            }

            _currentSystemEntity = null;
            _previosSystemEntity = null;
        }

        private void AddToConsistSystemPanelAction()
        {
            SS_SystemEntity startSystemEntity = ElemPicker(true, "Выбери элемент цепи, который нужно закольцевать на щите\\панели");

            //ConnectorSet mepConnectors = systemEntity.CurrentFamInst.MEPModel.ConnectorManager.Connectors;
            //Connector currentSysTypeConn = null;
            //foreach (Connector conn in mepConnectors)
            //{
            //    if (conn.ElectricalSystemType == viewEntity.SelectedSystemType && conn.AllRefs.Size == 0)
            //    {
            //        currentSysTypeConn = conn;
            //        break;
            //    }
            //}

            //if (currentSysTypeConn == null)
            //    throw new Exception($"У семейства ID: {systemEntity.CurrentFamInst.Id} - не осталось свободного соединителя нужноного типа");
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

        /// <summary>
        /// Команда для обновления цепи после внесения измов в цепь (без добавления или удаления элементов)
        /// </summary>
        private void RefreshSystemAction()
        {
            SS_SystemEntity startSystemEntity = ElemPicker(false, "Выбери ГОЛОВНОЙ элемент цепи, чтобы запустить анализ");
            if (startSystemEntity.CurrentElSystemSet == null || startSystemEntity.CurrentElSystemSet.Count() == 0)
            {
                TaskDialog taskDialog = new TaskDialog("KPLN: Внимание!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Элемент не подключен к цепи",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                taskDialog.Show();
            }
            else if (startSystemEntity.CurrentElSystemSet.Count() > 1)
            {
                TaskDialog taskDialog = new TaskDialog("KPLN: Внимание!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Выбран промежуточный элемент. Анализ невозможен.",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                taskDialog.Show();
            }
            else
            {
                startSystemEntity.SetSysParamsInstData();
                
                List<SS_SystemEntity> allEntitesFromSystem = new List<SS_SystemEntity>();
                
                // Берем данные по базовому щиту для запуска группирования по цепи
                FamilyInstance baseEquip = startSystemEntity.CurrentElSystemSet.FirstOrDefault().BaseEquipment;
                if (baseEquip != null)
                {
                    SS_SystemEntity startEntity = new SS_SystemEntity(baseEquip, _mainWindow.CurrentSystemViewEntity.SelectedSystemType).SetSysParamsInstData();
                    allEntitesFromSystem.Add(startEntity);

                    AddSystemEntityByBaseEquip(allEntitesFromSystem, startEntity);
                }
                else
                {
                    TaskDialog taskDialog = new TaskDialog("KPLN: Ошибка!")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                        MainContent = $"Не удалось определить стартовый щит для эл-та с id:{startSystemEntity.CurrentFamInst.Id}. Скинь разработчику!",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    taskDialog.Show();
                }

                if (allEntitesFromSystem.Any())
                {
                    using (Transaction trans = new Transaction(_doc, "KPLN: Обновление цепи"))
                    {
                        trans.Start();
                        
                        Readress(allEntitesFromSystem);
                        
                        trans.Commit();
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// Получить коллекцию SS_SystemEntity, которые объеденены в одну систему по спец. моделированию
        /// </summary>
        /// <param name="entitiesColl">Коллекция, которая будет содержать ВСЕ SS_SystemEntity</param>
        /// <param name="entity">SS_SystemEntity, с каторого начинается система</param>
        private void AddSystemEntityByBaseEquip(List<SS_SystemEntity> entitiesColl, SS_SystemEntity startEntity)
        {
            foreach(ElectricalSystem elSys in startEntity.CurrentElSystemSet)
            {
                ElementSet depElems = elSys.Elements;
                if (depElems.Size > 1)
                    throw new Exception($"Скинь разработчику: у цепи с id {elSys.Id} нарушена логика 1 участок цепи - 2 элемента.");

                SS_SystemEntity depEntity;
                IEnumerator depElemsEnum = depElems.GetEnumerator();
                while (depElemsEnum.MoveNext())
                {
                    if (depElemsEnum.Current is FamilyInstance famInts)
                    {
                        // Получаю экземпляр SS_SystemEntity подключаемого элемента
                        if (famInts.Id != startEntity.CurrentFamInst.Id && !entitiesColl.Select(ent => ent.CurrentFamInst.Id).Contains(famInts.Id))
                        {
                            depEntity = new SS_SystemEntity(famInts, _mainWindow.CurrentSystemViewEntity.SelectedSystemType).SetSysParamsInstData();
                            entitiesColl.Add(depEntity);

                            AddSystemEntityByBaseEquip(entitiesColl, depEntity);
                        }
                        // Запасной контроль наличия текущего эл-та в коллекции
                        else if (!entitiesColl.Select(ent => ent.CurrentFamInst.Id).Contains(famInts.Id))
                        {
                            depEntity = new SS_SystemEntity(famInts, _mainWindow.CurrentSystemViewEntity.SelectedSystemType).SetSysParamsInstData();
                            entitiesColl.Add(depEntity);
                        }
                    }
                    else
                        throw new Exception($"Для эл-та с id: {elSys.Id} - вы выбрали не экземпляр семейства, а что-то другое");
                }
            }
        }

        /// <summary>
        /// Переадресация внутри отсортированной ревит-цепи
        /// </summary>
        /// <param name="entitiesColl">Отсортированная цепь, от головы к хвосту</param>
        private void Readress(List<SS_SystemEntity> entitiesColl)
        {
            SS_SystemEntity frstEnt = entitiesColl.FirstOrDefault();
            int strNumber = frstEnt.CurrentAdressIndex - 1;

            SS_SystemEntity tempPrevEnt = null;
            foreach (SS_SystemEntity entity in entitiesColl)
            {
                if (tempPrevEnt != null)
                {
                    string prevAdress = tempPrevEnt
                        .CurrentFamInst
                        .LookupParameter(SS_SystemEntity.PreviousAdressParamName)
                        .AsValueString();

                    entity
                        .CurrentFamInst
                        .LookupParameter(SS_SystemEntity.PreviousAdressParamName)
                        .Set($"{prevAdress}");
                }
                
                strNumber += entity.CurrentAdressCountData;
                entity
                    .CurrentFamInst
                    .LookupParameter(SS_SystemEntity.CurrentAdressParamName)
                    .Set($"{entity.CurrentAdressSystemIndex}{entity.CurrentAdressSeparator}{strNumber}");
                
                tempPrevEnt = entity;
            }
        }

        /// <summary>
        /// Выбор элементов в проекте и генерация SS_SystemEntity
        /// </summary>
        private SS_SystemEntity ElemPicker(bool clearOldSys, string msg)
        {
            try
            {
                ElementId selectedId = _uidoc
                    .Selection
                    .PickObject(ObjectType.Element, new PickFilter(), msg)
                    .ElementId;

                Element elem = _doc.GetElement(selectedId);
                if (elem is FamilyInstance famInst)
                {
                    SS_SystemEntity resultSysEntity = new SS_SystemEntity(famInst, _mainWindow.CurrentSystemViewEntity.SelectedSystemType);

                    IEnumerable<ElectricalSystem> sysColl = resultSysEntity
                        .CurrentElSystemSet
                        .Where(sys => sys.SystemType.Equals(_mainWindow.CurrentSystemViewEntity.SelectedSystemType));
                    // Очистка от старой системы, если это нужно
                    if (clearOldSys && sysColl.Count() > 1)
                    {
                        string[] connPreviousAdressParamsData = sysColl
                            .Select(sys => sys.LookupParameter(SS_SystemEntity.PreviousAdressParamName).AsValueString())
                            .OrderBy(str => str)
                            .ToArray();
                        string[] connCurrentAdressParamsData = sysColl
                            .Select(sys => sys.LookupParameter(SS_SystemEntity.CurrentAdressParamName).AsValueString())
                            .OrderBy(str => str)
                            .ToArray();

                        SS_Sysetm_SelectIncludeElemAlg algForm = new SS_Sysetm_SelectIncludeElemAlg(
                            "Внимание!", 
                            "Элемент уже подключен к цепи с обеих сторон. Сейчас будет удалена цепь ПОСЛЕ текущего элемента. " +
                            "\nВНИМАНИЕ: если дальше этот элемент не подключить - он будет выступать в роли новой панели, т.е. цепь разделиться на 2 участка.");
                        algForm.AddCustomBtn($"{connPreviousAdressParamsData[1]}~{connCurrentAdressParamsData[1]}");
                        if ((bool)algForm.ShowDialog())
                        {
                            using (Transaction t = new Transaction(_doc, "KPLN: Удаление старой цепи"))
                            {
                                t.Start();
                                
                                _doc.Delete(resultSysEntity
                                    .CurrentElSystemSet
                                    .Where(sys => 
                                        $"{sys.LookupParameter(SS_SystemEntity.PreviousAdressParamName).AsValueString()}~{sys.LookupParameter(SS_SystemEntity.CurrentAdressParamName).AsValueString()}"
                                        .Equals(algForm.ClickedBtnContent))
                                    .FirstOrDefault()
                                    .Id);

                                t.Commit();
                            }
                        }
                        else
                            return null;
                    }

                    return resultSysEntity;
                }

                return null;
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
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSystemViewEntity;
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

            if (viewEntity.IsLineDraw == true && viewEntity.SelectedLineStyle == null)
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
        private void SetElementSystemData()
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSystemViewEntity;

            #region "КП_И_Адрес текущий"
            Parameter currentSysNameParam = _currentSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.CurrentAdressParamName) ??
                throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.CurrentAdressParamName}");

            Parameter currentElemPositionParam = _currentSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.PositionParamName) ??
                _currentSystemEntity.CurrentFamInst.Symbol.LookupParameter(SS_SystemEntity.PositionParamName) ??
                throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.PositionParamName}");

            string currentSysNameData = string.Format(
                "{0}{1}{2}{1}{3}",
                currentElemPositionParam.AsString(),
                viewEntity.UserSeparator,
                viewEntity.UserSystemIndex,
                viewEntity.StartNumber);
            currentSysNameParam.Set(currentSysNameData);
            #endregion

            #region "КП_И_Адрес предыдущий"
            if (_previosSystemEntity != null)
            {
                Parameter previousAdressParam = _currentSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.PreviousAdressParamName) ??
                    throw new Exception($"У элемента с id: {_currentSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.PreviousAdressParamName}");

                Parameter previousSysNameParam = _previosSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.CurrentAdressParamName) ??
                    throw new Exception($"У элемента с id: {_previosSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.CurrentAdressParamName}");

                previousAdressParam.Set(previousSysNameParam.AsString());
            }
            #endregion
        }

        /// <summary>
        /// Заполнение праматеров "КП_И_Адрес текущий", "КП_И_Адрес предыдущий", "КП_И_Префикс системы" для ЦЕПЕЙ
        /// </summary>
        /// <returns></returns>
        private bool SetSystemSystemData(ElectricalSystem elSys)
        {
            SS_SystemViewEntity viewEntity = _mainWindow.CurrentSystemViewEntity;

            #region "КП_И_Адрес текущий"
            Parameter currentSysNameParam = elSys.LookupParameter(SS_SystemEntity.CurrentAdressParamName) ??
                throw new Exception($"У цепи с id: {elSys.Id} нет параметра {SS_SystemEntity.CurrentAdressParamName}");

            Parameter currentElemPositionParam = _currentSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.PositionParamName) ??
                _currentSystemEntity.CurrentFamInst.Symbol.LookupParameter(SS_SystemEntity.PositionParamName) ??
                throw new Exception($"У элемена с id: {_currentSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.PositionParamName}");

            string currentSysNameData = string.Format(
                "{0}{1}{2}{1}{3}",
                currentElemPositionParam.AsString(), 
                viewEntity.UserSeparator, 
                viewEntity.UserSystemIndex, 
                viewEntity.StartNumber);
            currentSysNameParam.Set(currentSysNameData);
            #endregion

            #region "КП_И_Адрес предыдущий"
            Parameter previousAdressParam = elSys.LookupParameter(SS_SystemEntity.PreviousAdressParamName) ??
               throw new Exception($"У цепи с id: {elSys.Id} нет параметра {SS_SystemEntity.PreviousAdressParamName}");

            Parameter previousElemAdressParam = _previosSystemEntity.CurrentFamInst.LookupParameter(SS_SystemEntity.CurrentAdressParamName) ??
                throw new Exception($"У элемена с id: {_previosSystemEntity.CurrentFamInst.Id} нет параметра {SS_SystemEntity.CurrentAdressParamName}");

            previousAdressParam.Set(previousElemAdressParam.AsString());
            #endregion

            #region "КП_И_Префикс системы"
            Parameter sysPrefixParam = elSys.LookupParameter(SS_SystemEntity.SysPrefixParamName) ??
               throw new Exception($"У цепи с id: {elSys.Id} нет параметра {SS_SystemEntity.SysPrefixParamName}");

            sysPrefixParam.Set(_mainWindow.CurrentSystemViewEntity.UserSystemIndex);
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
                    detCurve.LineStyle = _mainWindow.CurrentSystemViewEntity.SelectedLineStyle.RevitGraphicsStyle;
                }
                else
                {
                    for (int i = 0; i < resultPnts.Count - 1; i++)
                    {
                        var line = Autodesk.Revit.DB.Line.CreateBound(resultPnts[i], resultPnts[i + 1]);
                        DetailCurve detCurve = _doc.Create.NewDetailCurve(_doc.ActiveView, line);
                        detCurve.LineStyle = _mainWindow.CurrentSystemViewEntity.SelectedLineStyle.RevitGraphicsStyle;
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
        private bool AreEqual(double coord1, double coord2, double epsilon) => Math.Abs(coord1 - coord2) < epsilon;
    }
}

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckFlatsArea : AbstrCheckCommand, IExternalCommand
    {
        #region Инициализация полей
        /// <summary>
        /// Коллекция параметров имени/номера (текст)
        /// </summary>
        private readonly List<RoomParamData> _roomNameParamDataColl = new List<RoomParamData>()
        {
            new RoomParamData("П_Имя", "Имя"),
            new RoomParamData("П_Номер", "Номер"),
        };

        /// <summary>
        /// Коллекция параметров площадей ВНУТРИ одной квартиры
        /// </summary>
        private readonly List<RoomParamData> _flatAreaParamDataColl = new List<RoomParamData>()
        {
            new RoomParamData("П_КВ_Площадь_Жилая", "КВ_Площадь_Жилая"),
            new RoomParamData("П_КВ_Площадь_Летние", "КВ_Площадь_Летние"),
            new RoomParamData("П_КВ_Площадь_Нежилая", "КВ_Площадь_Нежилая"),
            new RoomParamData("П_КВ_Площадь_Общая", "КВ_Площадь_Общая"),
            new RoomParamData("П_КВ_Площадь_Общая_К", "КВ_Площадь_Общая_К"),
            new RoomParamData("П_КВ_Площадь_Отапливаемые", "КВ_Площадь_Отапливаемые"),
        };

        /// <summary>
        /// Параметры площадей ВНУТРИ одной квартиры, которые необходимо проверять по СУММЕ на квартиру
        /// </summary>
        private readonly RoomParamData _flatAreaSumParamData = new RoomParamData("П_Площадь", "Площадь");

        /// <summary>
        /// Коллекция параметров площадей ОТДЕЛЬНОГО помещения
        /// </summary>
        private readonly List<RoomParamData> _roomAreaParamDataColl = new List<RoomParamData>()
        {
            new RoomParamData("П_Площадь", "Площадь"),
            new RoomParamData("П_ПОМ_Площадь", "ПОМ_Площадь"),
            new RoomParamData("П_ПОМ_Площадь_К", "ПОМ_Площадь_К"),
        };

        /// <summary>
        /// Коллекция параметров площадей ОТДЕЛЬНОГО помещения для СМЕЖНОЙ проверки
        /// </summary>
        private readonly List<RoomParamData> _roomAreaAdjacentParamDataColl = new List<RoomParamData>()
        {
            new RoomParamData("ПОМ_Площадь", "Площадь"),
        };
        #endregion

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            _name = "Проверка помещений";
            _application = uiapp;

            _allStorageName = "KPLN_CheckFlatsArea";
            _markerFieldName = "kpln_ar_area";

            _markerGuid = new Guid("720080C5-DA99-40D7-9445-E53F288AA149");
            _lastRunGuid = new Guid("720080C5-DA99-40D7-9445-E53F288AA150");
            _userTextGuid = new Guid("720080C5-DA99-40D7-9445-E53F288AA151");

            // Из-за сторонней библиотеки - нужно жестко (без общей абсатракции) прописать FieldName и StorageName у ExtensibleStorageBuilder
            _esBuildergMarker = new ExtensibleStorageBuilder(_markerGuid, _markerFieldName, "storage");

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            IEnumerable<Room> roomsColl = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble() > 0);

            #region Проверяю и обрабатываю элементы
            IEnumerable<WPFEntity> wpfColl = CheckCommandRunner(doc, roomsColl);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl, true);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion
            

            return Result.Failed;
        }

        private protected override List<CheckCommandError> CheckElements(Document doc, IEnumerable<Element> elemColl)
        {
            if (!(elemColl.Any()))
                throw new UserException("В проекте нет помещений.");

            foreach (Element elem in elemColl)
            {
                if (!(elem is Room room)) _criticalErrorElemColl.Add(new CheckCommandError(elem, "Не помещение!"));
                else
                {
                    List<RoomParamData> tempColl = new List<RoomParamData>(_roomNameParamDataColl);
                    tempColl.AddRange(_flatAreaParamDataColl);
                    tempColl.AddRange(_roomAreaParamDataColl);

                    foreach (RoomParamData rpc in tempColl)
                    {
                        if (rpc.FirstParam != null)
                            CheckParam(room, rpc.FirstParam);
                        if (rpc.SecondParam != null)
                            CheckParam(room, rpc.SecondParam);
                    }
                }
            }

            return null;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, IEnumerable<Element> elemColl)
        {
            List<WPFEntity> entities = new List<WPFEntity>();

            var roomDictTuple = PrepareRoomsDict(elemColl);

            Dictionary<string, List<Room>> flatRoomsDict = roomDictTuple.Item1;
            entities.AddRange(CheckFlatRoomsDataParams(flatRoomsDict));

            Dictionary<string, List<Room>> otherRoomsDict = roomDictTuple.Item2;
            entities.AddRange(CheckOtherRoomsDataParams(otherRoomsDict));

            return entities.OrderBy(e => e.ElementId.IntegerValue);
        }

        private void CheckParam(Room room, string paramName)
        {
            Parameter param = room.LookupParameter(paramName);
            if (param == null)
                throw new Exception($"У помещений нет параметра: {paramName}");
        }

        /// <summary>
        /// Подготовка помещений квартир [0] и всех остальных помещений [1]
        /// </summary>
        private (Dictionary<string, List<Room>>, Dictionary<string, List<Room>>) PrepareRoomsDict(IEnumerable<Element> elemsColl)
        {
            Dictionary<string, List<Room>> flatRoomResult = new Dictionary<string, List<Room>>();
            Dictionary<string, List<Room>> otherRoomResult = new Dictionary<string, List<Room>>();
            foreach (Room room in elemsColl)
            {
                // Утверждаю, что все квартиры имеют заполненный параметр "КВ_Номер"
                var paramData = room.LookupParameter("КВ_Номер").AsString();
                if (paramData != null && paramData != string.Empty)
                {
                    if (flatRoomResult.ContainsKey(paramData))
                        flatRoomResult[paramData].Add(room);
                    else
                        flatRoomResult.Add(paramData, new List<Room> { room });
                }

                else
                {
                    paramData = room.LookupParameter("Назначение").AsString();
                    if (paramData == null && paramData != string.Empty)
                        throw new Exception($"У помещения {room.Id} не заполнен параметр \"Назначение\". Это грубая ошибка, нужно исправить");

                    if (otherRoomResult.ContainsKey(paramData))
                        otherRoomResult[paramData].Add(room);
                    else
                        otherRoomResult.Add(paramData, new List<Room> { room });
                }
            }

            flatRoomResult = flatRoomResult.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
            otherRoomResult = otherRoomResult.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);

            return (flatRoomResult, otherRoomResult);
        }

        /// <summary>
        /// Набор проверок для квартир
        /// </summary>
        private List<WPFEntity> CheckFlatRoomsDataParams(Dictionary<string, List<Room>> roomsDict)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (KeyValuePair<string, List<Room>> kvp in roomsDict)
            {
                List<WPFEntity> entSumArea = EqualFlatSumAreaParamErrorData(kvp.Value);
                if (entSumArea != null)
                    result.AddRange(entSumArea);

                List<WPFEntity> entFlatArea = EqualFlatAreaParamErrorData(kvp.Value);
                if (entFlatArea != null)
                    result.AddRange(entFlatArea);

                List<WPFEntity> entArea = EqualRoomAreaParamErrorData(kvp.Value);
                if (entArea != null)
                    result.AddRange(entArea);

                List<WPFEntity> entName = EqualRoomNameParamErrorData(kvp.Value);
                if (entName != null)
                    result.AddRange(entName);

                List<WPFEntity> entAreaAdjacent = EqualRoomAreaAdjacentParamErrorData(kvp.Value);
                if (entAreaAdjacent != null)
                    result.AddRange(entAreaAdjacent);
            }

            return result;
        }

        /// <summary>
        /// Набор проверок для помещений, кроме квартир
        /// </summary>
        private List<WPFEntity> CheckOtherRoomsDataParams(Dictionary<string, List<Room>> roomsDict)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (KeyValuePair<string, List<Room>> kvp in roomsDict)
            {
                List<WPFEntity> entArea = EqualRoomAreaParamErrorData(kvp.Value);
                if (entArea != null)
                    result.AddRange(entArea);

                List<WPFEntity> entName = EqualRoomNameParamErrorData(kvp.Value);
                if (entName != null)
                    result.AddRange(entName);

                List<WPFEntity> entAreaAdjacent = EqualRoomAreaAdjacentParamErrorData(kvp.Value);
                if (entAreaAdjacent != null)
                    result.AddRange(entAreaAdjacent);
            }

            return result;
        }

        /// <summary>
        /// Поиск ошибок для помещений (в рамках допуска), у которых параметр должен определяться по сумме на квартиру
        /// </summary>
        private List<WPFEntity> EqualFlatSumAreaParamErrorData(List<Room> rooms)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            double fParamSumArea = rooms.Sum(r => r.LookupParameter(_flatAreaSumParamData.FirstParam).AsDouble());
            double sParamSumArea = rooms.Sum(r => r.LookupParameter(_flatAreaSumParamData.SecondParam).AsDouble());
            if (Math.Abs(fParamSumArea - sParamSumArea) > _flatAreaSumParamData.SumAreaTolerance)
            {
                foreach (Room room in rooms)
                {
                    Parameter fParam = room.LookupParameter(_flatAreaSumParamData.FirstParam);
                    Parameter sParam = room.LookupParameter(_flatAreaSumParamData.SecondParam);
                    if (fParam.AsDouble() != sParam.AsDouble())
                    {
                        Status currentStatus;
                        string approveComment = string.Empty;
                        if (ESBuilderUserText.IsDataExists_Text((Element)room)) 
                        {
                            currentStatus = Status.Approve;
                            approveComment = ESBuilderUserText.GetResMessage_Element((Element)room).Description;
                        }
                        else
                            currentStatus = Status.Error;
                        
                        result.Add(new WPFEntity(
                            room,
                            currentStatus,
                            "Нарушение суммарной площади (физической) квартиры",
                            $"Параметр \"{_flatAreaSumParamData.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                            false,
                            true,
                            $"Данные для сравнения со стадией П получены из параметра: \"{_flatAreaSumParamData.FirstParam}\".\nДопустимая разница внутри квартиры - 1 м²",
                            approveComment));
                    }
                }
            }

            if (result.Count > 0)
                return result;

            return null;
        }

        /// <summary>
        /// Поиск ошибок для помещений квартиры (в рамках допуска)
        /// </summary>
        private List<WPFEntity> EqualFlatAreaParamErrorData(List<Room> rooms)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            foreach (Room room in rooms)
            {
                foreach (RoomParamData fpc in _flatAreaParamDataColl)
                {
                    Parameter fParam = room.LookupParameter(fpc.FirstParam);
                    Parameter sParam = room.LookupParameter(fpc.SecondParam);
                    if (Math.Abs(fParam.AsDouble() - sParam.AsDouble()) > fpc.SumAreaTolerance)
                    {
                        if (fParam.AsDouble() != sParam.AsDouble())
                        {
                            Status currentStatus;
                            string approveComment = string.Empty;
                            if (ESBuilderUserText.IsDataExists_Text((Element)room))
                            {
                                currentStatus = Status.Approve;
                                approveComment = ESBuilderUserText.GetResMessage_Element((Element)room).Description;
                            }
                            else
                                currentStatus = Status.Error;

                            result.Add(new WPFEntity(
                                room,
                                currentStatus,
                                "Нарушение площади (в марках) квартиры",
                                $"Параметр \"{fpc.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                                false,
                                true,
                                $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\"." +
                                    $"\nДопустимая разница внутри помещения - 1 м².\nВыделено только ОДНО помещение квартиры",
                                approveComment));
                            continue;
                        }

                    }
                }

            }

            if (result.Count > 0)
                return result;

            return null;
        }

        /// <summary>
        /// Поиск ошибок для имени и номера помещений
        /// </summary>
        private List<WPFEntity> EqualRoomNameParamErrorData(List<Room> rooms)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            foreach (Room room in rooms)
            {
                foreach (RoomParamData fpc in _roomNameParamDataColl)
                {
                    Parameter fParam = room.LookupParameter(fpc.FirstParam);
                    Parameter sParam = room.LookupParameter(fpc.SecondParam);
                    string fParamToString = fParam.AsString();
                    string sParamToString = sParam.AsString();
                    // Проверка на текстовый тип данных (имена помещений)
                    if (fParamToString == null || sParamToString == null)
                    {
                        fParamToString = fParam.AsValueString();
                        sParamToString = sParam.AsValueString();

                        if (fParamToString == null || sParamToString == null)
                        {
                            result.Add(new WPFEntity(
                                room,
                                SetApproveStatusByUserComment(room, Status.Error),
                                "Нарушение анализа данных",
                                $"Помещение было создано после фиксации площадей на стадии П. Необходимо согласовать добавление с ГАПом и выполнить процесс фиксации (ТОЛЬКО через BIM-отдел)",
                                false,
                                true,
                                null,
                                GetUserComment(room)));
                            break;
                        }
                    }

                    else if (!fParamToString.Equals(sParamToString))
                    {
                        result.Add(new WPFEntity(
                            room,
                            SetApproveStatusByUserComment(room, Status.Error),
                            "Нарушение имени/номера помещения",
                            $"Параметр \"{fpc.SecondParam}\" на стадии П был \"{fParamToString}\", сейчас - \"{sParamToString}\"",
                            false,
                            true,
                            $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\".",
                            GetUserComment(room)));
                    }
                }
            }

            if (result.Count > 0)
                return result;

            return null;
        }

        /// <summary>
        /// Поиск ошибок для помещений, у которых параметр отличается от исходного в рамках допуска
        /// </summary>
        private List<WPFEntity> EqualRoomAreaParamErrorData(List<Room> rooms)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            foreach (Room room in rooms)
            {
                foreach (RoomParamData fpc in _roomAreaParamDataColl)
                {
                    Parameter fParam = room.LookupParameter(fpc.FirstParam);
                    Parameter sParam = room.LookupParameter(fpc.SecondParam);
                    if (Math.Abs(fParam.AsDouble() - sParam.AsDouble()) > fpc.SumAreaTolerance)
                    {
                        if (fParam.AsDouble() != sParam.AsDouble())
                        {
                            Status currentStatus;
                            string approveComment = string.Empty;
                            if (ESBuilderUserText.IsDataExists_Text((Element)room))
                            {
                                currentStatus = Status.Approve;
                                approveComment = ESBuilderUserText.GetResMessage_Element((Element)room).Description;
                            }
                            else
                                currentStatus = Status.Error;

                            result.Add(new WPFEntity(
                                room,
                                currentStatus,
                                "Нарушение площади отдельного помещения",
                                $"Параметр \"{fpc.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                                false,
                                true,
                                $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\"." +
                                    $"\nДопустимая разница внутри помещения - 1 м²",
                                approveComment));
                        }

                    }
                }

            }

            if (result.Count > 0)
                return result;

            return null;
        }

        /// <summary>
        /// Поиск ошибок для помещений, у которых смежные параметры отличаются в рамках допуска
        /// </summary>
        private List<WPFEntity> EqualRoomAreaAdjacentParamErrorData(List<Room> rooms)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            foreach (Room room in rooms)
            {
                foreach (RoomParamData fpc in _roomAreaAdjacentParamDataColl)
                {
                    Parameter fParam = room.LookupParameter(fpc.FirstParam);
                    Parameter sParam = room.LookupParameter(fpc.SecondParam);
                    if (Math.Abs(fParam.AsDouble() - sParam.AsDouble()) > 1.0)
                    {
                        if (fParam.AsDouble() != sParam.AsDouble())
                        {
                            Status currentStatus;
                            string approveComment = string.Empty;
                            if (ESBuilderUserText.IsDataExists_Text((Element)room))
                            {
                                currentStatus = Status.Approve;
                                approveComment = ESBuilderUserText.GetResMessage_Element((Element)room).Description;
                            }
                            else
                                currentStatus = Status.Error;

                            result.Add(new WPFEntity(
                                room,
                                currentStatus,
                                "Квартирография не запускалась",
                                $"Рельаная площадь помещения (\"{fpc.SecondParam}\") {sParam.AsValueString()}, а в марке (\"{fpc.FirstParam}\") - {fParam.AsValueString()}",
                                false,
                                true,
                                $"По согласованию с ГАПом - необходимо запустить квартирографию",
                                approveComment));
                        }

                    }
                }

            }

            if (result.Count > 0)
                return result;

            return null;
        }

        /// <summary>
        ///  Расчте расстояния Дамерлоу-Левинштейна для текста
        /// </summary>
        private static int DamerauLevenshteinDistance(string firstText, string secondText)
        {
            var n = firstText.Length + 1;
            var m = secondText.Length + 1;
            var arrayD = new int[n, m];

            for (var i = 0; i < n; i++)
            {
                arrayD[i, 0] = i;
            }

            for (var j = 0; j < m; j++)
            {
                arrayD[0, j] = j;
            }

            for (var i = 1; i < n; i++)
            {
                for (var j = 1; j < m; j++)
                {
                    var cost = firstText[i - 1] == secondText[j - 1] ? 0 : 1;

                    arrayD[i, j] = Minimum(arrayD[i - 1, j] + 1,          // удаление
                                            arrayD[i, j - 1] + 1,         // вставка
                                            arrayD[i - 1, j - 1] + cost); // замена

                    if (i > 1 && j > 1
                        && firstText[i - 1] == secondText[j - 2]
                        && firstText[i - 2] == secondText[j - 1])
                    {
                        arrayD[i, j] = Minimum(arrayD[i, j],
                                           arrayD[i - 2, j - 2] + cost); // перестановка
                    }
                }
            }

            return arrayD[n - 1, m - 1];
        }

        private static int Minimum(int a, int b) => a < b ? a : b;

        private static int Minimum(int a, int b, int c) => (a = a < b ? a : b) < c ? a : c;
    }
}

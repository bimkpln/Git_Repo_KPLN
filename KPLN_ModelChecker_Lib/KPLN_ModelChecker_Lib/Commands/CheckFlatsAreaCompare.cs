using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    /// <summary>
    /// Контейнер данных по помещению
    /// </summary>
    internal sealed class RoomParamData
    {
        /// <summary>
        /// Допуск в 1 м² (double - в футах!) в перделах 1 квартиры
        /// </summary>
        internal readonly double SumAreaTolerance = 10.764;

        /// <summary>
        /// Допуск в N символов для разницы в именовании и нумерации
        /// </summary>
        internal readonly int TextTolerance = 1;

        internal RoomParamData(string firstParamName, string secondParamName)
        {
            FirstParam = firstParamName;
            SecondParam = secondParamName;
        }

        internal string FirstParam { get; private set; }

        internal string SecondParam { get; private set; }
    }

    public sealed class CheckFlatsAreaCompare : AbstrCheck
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

        public CheckFlatsAreaCompare() : base()
        {
            if (PluginName == null)
                PluginName = "АР_Р: Сравнение помещений";

            if (ESEntity == null)
            {
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckFlatsArea",
                    new Guid("720080C5-DA99-40D7-9445-E53F288AA150"),
                    new Guid("720080C5-DA99-40D7-9445-E53F288AA151"),
                    new Guid("720080C5-DA99-40D7-9445-E53F288AA155"));

                // Из-за сторонней библиотеки (на python) - нужно жестко (без общей абсатракции) прописать FieldName и StorageName у ExtensibleStorageBuilder
                ESEntity.ESBuildergMarker = new ExtensibleStorageBuilder(ESEntity.MarkerGuid, "Last_Run", "KPLN_ARArea");
            }
        }

        public override Element[] GetElemsToCheck() => new FilteredElementCollector(CheckDocument)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Cast<Room>()
            .ToArray();

        private protected override void CheckRElems_SetElemErrorColl(object[] objColl)
        {
            if (!(objColl.Any()))
                throw new CheckerException("В проекте нет помещений.");

            foreach (object obj in objColl)
            {
                if (obj is Element element)
                {
                    if (!(element is Room room))
                        PrepareElemsErrorColl.Append(new CheckCommandMsg(element, "Не помещение!"));
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
                else throw new Exception("Ошибка анализируемой коллекции. Отправь разработчику");
            }
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            var roomDictTuple = PrepareRoomsDict(elemColl);

            Dictionary<string, List<Room>> flatRoomsDict = roomDictTuple.Item1;
            _checkerEntitiesCollHeap.AddRange(CheckFlatRoomsDataParams(flatRoomsDict));

            Dictionary<string, List<Room>> otherRoomsDict = roomDictTuple.Item2;
            _checkerEntitiesCollHeap.AddRange(CheckOtherRoomsDataParams(otherRoomsDict));

            return CheckResultStatus.Succeeded;
        }

        private void CheckParam(Room room, string paramName)
        {
            _ = room.LookupParameter(paramName) ?? throw new CheckerException($"У помещений нет параметра: {paramName}");
        }

        /// <summary>
        /// Подготовка помещений квартир [0] и всех остальных помещений [1]
        /// </summary>
        private (Dictionary<string, List<Room>>, Dictionary<string, List<Room>>) PrepareRoomsDict(IEnumerable<Element> elemsColl)
        {
            Dictionary<string, List<Room>> flatRoomResult = new Dictionary<string, List<Room>>();
            Dictionary<string, List<Room>> otherRoomResult = new Dictionary<string, List<Room>>();
            foreach (Room room in elemsColl.Cast<Room>())
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
        private List<CheckerEntity> CheckFlatRoomsDataParams(Dictionary<string, List<Room>> roomsDict)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (KeyValuePair<string, List<Room>> kvp in roomsDict)
            {
                List<CheckerEntity> entSumArea = EqualFlatSumAreaParamErrorData(kvp.Value);
                if (entSumArea != null)
                    result.AddRange(entSumArea);

                List<CheckerEntity> entFlatArea = EqualFlatAreaParamErrorData(kvp.Value);
                if (entFlatArea != null)
                    result.AddRange(entFlatArea);

                List<CheckerEntity> entArea = EqualRoomAreaParamErrorData(kvp.Value);
                if (entArea != null)
                    result.AddRange(entArea);

                List<CheckerEntity> entName = EqualRoomNameParamErrorData(kvp.Value);
                if (entName != null)
                    result.AddRange(entName);

                List<CheckerEntity> entAreaAdjacent = EqualRoomAreaAdjacentParamErrorData(kvp.Value);
                if (entAreaAdjacent != null)
                    result.AddRange(entAreaAdjacent);
            }

            return result;
        }

        /// <summary>
        /// Набор проверок для помещений, кроме квартир
        /// </summary>
        private List<CheckerEntity> CheckOtherRoomsDataParams(Dictionary<string, List<Room>> roomsDict)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (KeyValuePair<string, List<Room>> kvp in roomsDict)
            {
                List<CheckerEntity> entArea = EqualRoomAreaParamErrorData(kvp.Value);
                if (entArea != null)
                    result.AddRange(entArea);

                List<CheckerEntity> entName = EqualRoomNameParamErrorData(kvp.Value);
                if (entName != null)
                    result.AddRange(entName);

                List<CheckerEntity> entAreaAdjacent = EqualRoomAreaAdjacentParamErrorData(kvp.Value);
                if (entAreaAdjacent != null)
                    result.AddRange(entAreaAdjacent);
            }

            return result;
        }

        /// <summary>
        /// Поиск ошибок для помещений (в рамках допуска), у которых параметр должен определяться по сумме на квартиру
        /// </summary>
        private List<CheckerEntity> EqualFlatSumAreaParamErrorData(List<Room> rooms)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

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
                        result.Add(new CheckerEntity(
                            room,
                            "Нарушение суммарной площади (физической) квартиры",
                            $"Параметр \"{_flatAreaSumParamData.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                            $"Данные для сравнения со стадией П получены из параметра: \"{_flatAreaSumParamData.FirstParam}\".\nДопустимая разница внутри квартиры - 1 м²",
                            false)
                            .Set_CanApprovedAndESData(ESEntity));
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
        private List<CheckerEntity> EqualFlatAreaParamErrorData(List<Room> rooms)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

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
                            result.Add(new CheckerEntity(
                                room,
                                "Нарушение площади (в марках) квартиры",
                                $"Параметр \"{fpc.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                                $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\"." +
                                    $"\nДопустимая разница внутри помещения - 1 м².\nВыделено только ОДНО помещение квартиры",
                                false)
                                .Set_CanApprovedAndESData(ESEntity));

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
        private List<CheckerEntity> EqualRoomNameParamErrorData(List<Room> rooms)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

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
                            result.Add(new CheckerEntity(
                                room,
                                "Нарушение анализа данных",
                                "Помещение было создано после фиксации площадей на стадии П. Необходимо согласовать добавление с ГАПом и выполнить процесс фиксации (ТОЛЬКО через BIM-отдел)",
                                string.Empty,
                                false)
                                .Set_CanApprovedAndESData(ESEntity));

                            break;
                        }
                    }

                    else if (!fParamToString.Equals(sParamToString))
                    {
                        result.Add(new CheckerEntity(
                            room,
                            "Нарушение имени/номера помещения",
                            $"Параметр \"{fpc.SecondParam}\" на стадии П был \"{fParamToString}\", сейчас - \"{sParamToString}\"",
                            $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\".",
                            false)
                            .Set_CanApprovedAndESData(ESEntity));
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
        private List<CheckerEntity> EqualRoomAreaParamErrorData(List<Room> rooms)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

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
                            result.Add(new CheckerEntity(
                                room,
                                "Нарушение площади отдельного помещения",
                                $"Параметр \"{fpc.SecondParam}\" на стадии П был {fParam.AsValueString()}, сейчас - {sParam.AsValueString()}",
                                $"Данные для сравнения со стадией П получены из параметра: \"{fpc.FirstParam}\"." +
                                    $"\nДопустимая разница внутри помещения - 1 м²",
                                false)
                                .Set_CanApprovedAndESData(ESEntity));
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
        private List<CheckerEntity> EqualRoomAreaAdjacentParamErrorData(List<Room> rooms)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

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
                            result.Add(new CheckerEntity(
                                room,
                                "Квартирография не запускалась",
                                $"Рельаная площадь помещения (\"{fpc.SecondParam}\") {sParam.AsValueString()}, а в марке (\"{fpc.FirstParam}\") - {fParam.AsValueString()}",
                                $"По согласованию с ГАПом - необходимо запустить квартирографию",
                                false)
                                .Set_CanApprovedAndESData(ESEntity));
                        }

                    }
                }

            }

            if (result.Count > 0)
                return result;

            return null;
        }
    }
}

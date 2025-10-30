using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools.Forms.Models.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.AR_PyatnGraph
{
    public enum RoomType
    {
        Other,
        Flat,
        Balkon,
        Loggia,
        Terrace
    }

    public struct LittleRoomData
    {
        public Element LRDElement { get; set; }
        public RoomType LRDRoomType { get; set; }
    }

    /// <summary>
    /// Контейнер данных из модели по ОТДЕЛЬНЫМ помещениям
    /// </summary>
    public sealed class ARPG_Room
    {
        private static readonly string _roomName = "Квартира";
        private static readonly string _balkName = "Балкон";
        private static readonly string _logName = "Лоджия";
        private static readonly string _terName = "Терраса";

        private static readonly string[] _roomNames = new string[]
        {
            _roomName,
            _balkName,
            _logName,
            _terName,
        };

        private ARPG_Room() { }

        /// <summary>
        /// Словарь ОШИБОК элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        public static Dictionary<string, List<ElementId>> ErrorDict_Room;

        /// <summary>
        /// Ссылка на элемент ревит
        /// </summary>
        public Element Elem_Room { get; private set; }

        /// <summary>
        /// Площадь квартиры из модели
        /// </summary>
        public double AreaData_Room { get; private set; }

        /// <summary>
        /// Имя квартиры\балкона\лоджии\террасы
        /// </summary>
        public string FlatNameData_Room { get; private set; }

        /// <summary>
        /// Номер квартиры
        /// </summary>
        public string FlatNumbData_Room { get; private set; }

        /// <summary>
        /// Номер этажа квартиры
        /// </summary>
        public string FlatLvlNumbData_Room { get; private set; }

        /// <summary>
        /// Захватка 1 (напр: корпус)
        /// </summary>
        public string GripCorpData_Room { get; private set; }

        /// <summary>
        /// Захватка 2 (напр: секция)
        /// </summary>
        public string GripSectData_Room { get; private set; }

        /// <summary>
        /// Значение допуска в м2
        /// </summary>
        public double FlatAreaTolerance_Room { get; private set; }

        /// <summary>
        /// Значение коэффициента
        /// </summary>
        public double AreaCoeff_Room { get; private set; }

        /// <summary>
        /// Площадь с коэффициентом
        /// </summary>
        public double AreaCoeffData_Room { get; private set; }

        /// <summary>
        /// Параметр для площади с коэффициентом
        /// </summary>
        public Parameter AreaCoeffParam_Room { get; private set; }

        /// <summary>
        /// Параметр для площади с коэффициентом
        /// </summary>
        public Parameter SumAreaCoeffParam_Room { get; private set; }

        /// <summary>
        /// Параметр для кода квартиры по ТЗ
        /// </summary>
        public Parameter TZCodeParam_Room { get; private set; }

        /// <summary>
        /// Параметр для имя диапазона по ТЗ
        /// </summary>
        public Parameter TZRangeNameParam_Room { get; private set; }

        /// <summary>
        /// Параметр для процент диапазона по ТЗ 
        /// </summary>
        public Parameter TZPercentParam_Room { get; private set; }

        /// <summary>
        /// Параметр для диапазон min по ТЗ
        /// </summary>
        public Parameter TZAreaMinParam_Room { get; private set; }

        /// <summary>
        /// Параметр для диапазон max по ТЗ
        /// </summary>
        public Parameter TZAreaMaxParam_Room { get; private set; }

        /// <summary>
        /// Параметр для полученного процента в модели
        /// </summary>
        public Parameter ModelPercentParam_Room { get; private set; }

        /// <summary>
        /// Параметр для полученного отклонения от процента из ТЗ в модели
        /// </summary>
        public Parameter ModelPercentToleranceParam_Room { get; private set; }

        /// <summary>
        /// Тип помещения
        /// </summary>
        public RoomType ARPG_RoomType { get; private set; }

        /// <summary>
        /// Получить коллекцию помещений в зависимости от анализируемого типа
        /// </summary>
        internal static ARPG_Room[] Get_ARPG_Rooms(Document doc, ARPG_TZ_MainData tzData, ARPG_DesOptEntity arpg_desOpt, bool presetFlatCode)
        {
            if (tzData.FlatType == nameof(FilledRegion))
            {
                tzData.FlatAreaParam = BuiltInParameter.ROOM_AREA;
                return GetFromFilledRegions(doc, tzData, arpg_desOpt);
            }
            else if (tzData.FlatType == nameof(Room))
            {
                tzData.FlatAreaParam = BuiltInParameter.ROOM_AREA;
                return GetFromRooms(doc, tzData, arpg_desOpt, presetFlatCode);
            }
            else
                throw new NotSupportedException(
                    string.Format("Тип элементов \"{0}\" не поддерживается.", tzData.FlatType));
        }

        /// <summary>
        /// Задать расчётные вторичные параметры
        /// </summary>
        internal static void SetCountedRoomData(ARPG_Room[] arpgRooms)
        {
            ErrorDict_Room = new Dictionary<string, List<ElementId>>();

            double arpgRooms_SummArea = arpgRooms.Sum(ar => ar.AreaCoeffData_Room);

            // Заполняю значения
            foreach (ARPG_Room arpgRoom in arpgRooms)
            {
                // Расчёт факт % и значения отклонения % от ТЗ
                ARPG_Room[] equalArpgRooms = arpgRooms
                    .Where(af =>
                        af.TZCodeParam_Room.AsString().Equals(arpgRoom.TZCodeParam_Room.AsString())
                        && af.TZRangeNameParam_Room.AsString().Equals(arpgRoom.TZRangeNameParam_Room.AsString()))
                    .ToArray();

                // Только одно помещение такого типа
                if (equalArpgRooms == null || equalArpgRooms.Length == 0)
                {
                    double factPercent = (arpgRoom.AreaCoeffData_Room / arpgRooms_SummArea) * 100;
                    SetCountedRoomData_Percents(arpgRoom, factPercent);
                }
                else
                {
                    double sumAreaCoeffDataEqualsARPGRooms = equalArpgRooms.Sum(ar => ar.AreaCoeffData_Room);
                    double factPercent = (sumAreaCoeffDataEqualsARPGRooms / arpgRooms_SummArea) * 100;
                    foreach (ARPG_Room equalArpgRoom in equalArpgRooms)
                    {
                        SetCountedRoomData_Percents(equalArpgRoom, factPercent);
                    }
                }
            }
        }

        private static void SetCountedRoomData_Percents(ARPG_Room arpgRoom, double factPercent)
        {
            double tzPercent = arpgRoom.TZPercentParam_Room.AsDouble();
            double factTolerance = factPercent - tzPercent;

            arpgRoom.ModelPercentParam_Room.Set(factPercent);
            arpgRoom.ModelPercentToleranceParam_Room.Set(factTolerance);
        }

        #region Цветовая область
        private static ARPG_Room[] GetFromFilledRegions(Document doc, ARPG_TZ_MainData tzData, ARPG_DesOptEntity arpg_desOpt)
        {
            ErrorDict_Room = new Dictionary<string, List<ElementId>>();

            FilledRegion[] filledRegions = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegion))
                .WhereElementIsNotElementType()
                .OfType<FilledRegion>()
                .Where(fr =>
                {
                    Parameter typeParam = fr.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);

                    bool isSelectedDO = fr.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID).AsValueString().Equals(arpg_desOpt.ARPG_DesignOptionId.ToString());
                    
                    return typeParam != null && typeParam.AsValueString() != null && typeParam.AsValueString().StartsWith("ТЭП_") && isSelectedDO;
                })
                .ToArray();

            // TODO: Реализовать при готовности тестовой модели
            return new ARPG_Room[0];
        }
        #endregion

        #region Помещения
        private static ARPG_Room[] GetFromRooms(Document doc, ARPG_TZ_MainData tzData, ARPG_DesOptEntity arpg_desOpt, bool presetFlatCode)
        {
            ErrorDict_Room = new Dictionary<string, List<ElementId>>();

            Room[] rooms = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType()
                .OfType<Room>()
                .Where(r =>
                {
                    bool validArea = r.Area > 0;
                    bool isSelectedDO = r.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID).AsValueString().Equals(arpg_desOpt.ARPG_DesignOptionId.ToString());

                    return validArea && isSelectedDO;
                })
                .ToArray();

            List<LittleRoomData> lrdColl = new List<LittleRoomData>();
            foreach (Room r in rooms)
            {
                RoomType rt = GetRoomType(tzData, r);
                if (rt == RoomType.Other)
                    continue;

                lrdColl.Add(new LittleRoomData()
                {
                    LRDElement = r,
                    LRDRoomType = rt,
                });
            }

            if (ErrorDict_Room.Any())
                return null;

            return Create_ARPGRooms(lrdColl, tzData, presetFlatCode);
        }

        private static ARPG_Room[] Create_ARPGRooms(IEnumerable<LittleRoomData> lrdColl, ARPG_TZ_MainData tzData, bool presetFlatCode)
        {
            List<ARPG_Room> result = new List<ARPG_Room>();

            bool isNoHeatingRoom = lrdColl.Any(lrd => lrd.LRDRoomType != RoomType.Flat);
            foreach (LittleRoomData lrd in lrdColl)
            {
                Element elem = lrd.LRDElement;
                if (lrd.LRDRoomType == RoomType.Other)
                {
                    HtmlOutput.SetMsgDict_ByMsg($"Отправь разработчику - не удалось определить помещение", elem.Id, ErrorDict_Room);
                    continue;
                }

                List<Parameter> paramsToCheck = new List<Parameter>();

                Parameter flatNumbParam = null;
                Parameter flatLvlNumbParam = null;
                Parameter gripCorpParam = null;
                Parameter gripSectParam = null;

                // Забираю доп. параметры 
                if (isNoHeatingRoom)
                {
                    flatNumbParam = elem.LookupParameter(tzData.FlatNumbParamName);
                    paramsToCheck.Add(flatNumbParam);
                    flatLvlNumbParam = elem.LookupParameter(tzData.FlatLvlNumbParamName);
                    paramsToCheck.Add(flatLvlNumbParam);
                    gripCorpParam = elem.LookupParameter(tzData.GripCorpParamName);
                    paramsToCheck.Add(gripCorpParam);
                    if (!string.IsNullOrEmpty(tzData.GripSectParamName))
                    {
                        gripSectParam = elem.LookupParameter(tzData.GripSectParamName);
                        paramsToCheck.Add(gripSectParam);
                    }

                    // Проверка на пропущенные параметры
                    CheckMissingParam(elem, flatNumbParam, tzData.FlatNumbParamName);
                    CheckMissingParam(elem, flatLvlNumbParam, tzData.FlatLvlNumbParamName);
                    CheckMissingParam(elem, gripCorpParam, tzData.GripCorpParamName);
                    if (gripSectParam != null)
                        CheckMissingParam(elem, gripSectParam, tzData.GripCorpParamName);
                }

                // Забираю основные параметры
                Parameter flatNameParam = elem.LookupParameter(tzData.FlatNameParamName);
                Parameter areaParam = elem.get_Parameter(BuiltInParameter.ROOM_AREA);
                paramsToCheck.Add(areaParam);
                Parameter areaCoeffParam = elem.LookupParameter(tzData.AreaCoeffParamName);
                paramsToCheck.Add(areaCoeffParam);
                Parameter sumAreaCoeffParam = elem.LookupParameter(tzData.SumAreaCoeffParamName);
                paramsToCheck.Add(sumAreaCoeffParam);
                Parameter tzCodeNameParam = elem.LookupParameter(tzData.TZCodeParamName);
                paramsToCheck.Add(tzCodeNameParam);
                Parameter tzRangeNameParam = elem.LookupParameter(tzData.TZRangeNameParamName);
                paramsToCheck.Add(tzRangeNameParam);
                Parameter tzPercentParam = elem.LookupParameter(tzData.TZPercentParamName);
                paramsToCheck.Add(tzPercentParam);
                Parameter tzAreaMinParam = elem.LookupParameter(tzData.TZAreaMinParamName);
                paramsToCheck.Add(tzAreaMinParam);
                Parameter tzAreaMaxParam = elem.LookupParameter(tzData.TZAreaMaxParamName);
                paramsToCheck.Add(tzAreaMaxParam);
                Parameter modelPercentParamName = elem.LookupParameter(tzData.ModelPercentParamName);
                paramsToCheck.Add(modelPercentParamName);
                Parameter modelPercentToleranceParamName = elem.LookupParameter(tzData.ModelPercentToleranceParamName);
                paramsToCheck.Add(modelPercentToleranceParamName);


                // Проверка на пропущенные параметры
                CheckMissingParam(elem, flatNameParam, tzData.FlatNameParamName);
                CheckMissingParam(elem, areaParam, elem.get_Parameter(tzData.FlatAreaParam).Definition.Name);
                CheckMissingParam(elem, areaCoeffParam, tzData.AreaCoeffParamName);
                CheckMissingParam(elem, sumAreaCoeffParam, tzData.SumAreaCoeffParamName);
                CheckMissingParam(elem, tzCodeNameParam, tzData.TZCodeParamName);
                CheckMissingParam(elem, tzRangeNameParam, tzData.TZRangeNameParamName);
                CheckMissingParam(elem, tzPercentParam, tzData.TZPercentParamName);
                CheckMissingParam(elem, tzAreaMinParam, tzData.TZAreaMinParamName);
                CheckMissingParam(elem, tzAreaMaxParam, tzData.TZAreaMaxParamName);
                CheckMissingParam(elem, modelPercentParamName, tzData.ModelPercentParamName);
                CheckMissingParam(elem, modelPercentToleranceParamName, tzData.ModelPercentToleranceParamName);

                // Общая сумма по проверяемым параметрам
                if (HasMissing(paramsToCheck))
                    continue;

                // Получаю значения
                string flatNameData = GetParamValue(flatNameParam);
                double areaData = areaParam.AsDouble();
                string flatNumbData = GetParamValue(flatNumbParam);
                string flatLvlNumbData = GetParamValue(flatLvlNumbParam);
                string gripCorpData = GetParamValue(gripCorpParam);
                string gripSectData = GetParamValue(gripSectParam);
                string tzCodeNameData = GetParamValue(tzCodeNameParam);

                // Проверка пустых значений
                CheckEmptyValue(elem, tzData.FlatNameParamName, flatNameData);
                if (!presetFlatCode)
                    CheckEmptyValue(elem, tzData.TZCodeParamName, tzCodeNameData);

                if (isNoHeatingRoom)
                {
                    CheckEmptyValue(elem, tzData.FlatNumbParamName, flatNumbData);
                    CheckEmptyValue(elem, tzData.FlatLvlNumbParamName, flatLvlNumbData);
                    
                    if (tzData.IsGripCorpParam && tzData.IsGripSectParam)
                    {
                        CheckEmptyValue(elem, tzData.GripCorpParamName, gripCorpData);
                        CheckEmptyValue(elem, tzData.GripSectParamName, gripSectData);
                    }
                    else if (tzData.IsGripCorpParam)
                        CheckEmptyValue(elem, tzData.GripCorpParamName, gripCorpData);
                    else if (tzData.IsGripSectParam)
                        CheckEmptyValue(elem, tzData.GripSectParamName, gripSectData);
                }

                if (ErrorDict_Room.Values.Any(v => v.Contains(elem.Id)))
                    continue;

                ARPG_Room aRPG_Room = new ARPG_Room()
                {
                    Elem_Room = elem,
                    AreaData_Room = areaData,
                    ARPG_RoomType = lrd.LRDRoomType,
                    FlatNameData_Room = flatNameData,
                    FlatNumbData_Room = flatNumbData,
                    FlatLvlNumbData_Room = flatLvlNumbData,
                    GripCorpData_Room = gripCorpData,
                    GripSectData_Room = gripSectData,
                    AreaCoeffParam_Room = areaCoeffParam,
                    SumAreaCoeffParam_Room = sumAreaCoeffParam,
                    TZCodeParam_Room = tzCodeNameParam,
                    TZRangeNameParam_Room = tzRangeNameParam,
                    TZPercentParam_Room = tzPercentParam,
                    TZAreaMinParam_Room = tzAreaMinParam,
                    TZAreaMaxParam_Room = tzAreaMaxParam,
                    ModelPercentParam_Room = modelPercentParamName,
                    ModelPercentToleranceParam_Room = modelPercentToleranceParamName,
                };

                aRPG_Room.Set_AreaCoeffAndFlatArea(tzData);

                result.Add(aRPG_Room);
            }

            return result.ToArray(); ;
        }


        //private static ARPG_Room[] Create_ARPGRooms1(IEnumerable<Element> elemsToCreate, ARPG_TZ_MainData tzData, bool presetFlatCode)
        //{
        //    List<ARPG_Room> result = new List<ARPG_Room>();
        //    foreach (Element elem in elemsToCreate)
        //    {
        //        if (tzData.FlatsFilter)
        //        {
        //            RoomType roomRoomType = GetRoomType(tzData, elem, true);
        //            if (roomRoomType == RoomType.Other)
        //                continue;

        //            List<Parameter> paramsToCheck = new List<Parameter>();
        //            // Если в проекте есть лоджии/балконы/террасы - нужно делать биндинг по помещениям
        //            Parameter balkTerLogNameParam = null;
        //            Parameter flatNumbParam = null;
        //            Parameter flatLvlNumbParam = null;
        //            Parameter gripParam1 = null;
        //            Parameter gripParam2 = null;

        //            // Забираю доп. параметры 
        //            if (tzData.HeatingRoomsInPrj)
        //            {
        //                balkTerLogNameParam = elem.LookupParameter(tzData.BalkTerLogNameParamName);
        //                paramsToCheck.Add(balkTerLogNameParam);
        //                flatNumbParam = elem.LookupParameter(tzData.FlatNumbParamName);
        //                paramsToCheck.Add(flatNumbParam);
        //                flatLvlNumbParam = elem.LookupParameter(tzData.FlatLvlNumbParamName);
        //                paramsToCheck.Add(flatLvlNumbParam);
        //                gripParam1 = elem.LookupParameter(tzData.GripParamName1);
        //                paramsToCheck.Add(gripParam1);
        //                if (!string.IsNullOrEmpty(tzData.GripParamName2))
        //                {
        //                    gripParam2 = elem.LookupParameter(tzData.GripParamName2);
        //                    paramsToCheck.Add(gripParam2);
        //                }

        //                // Проверка на пропущенные параметры
        //                CheckMissingParam(elem, balkTerLogNameParam, tzData.BalkTerLogNameParamName);
        //                CheckMissingParam(elem, flatNumbParam, tzData.FlatNumbParamName);
        //                CheckMissingParam(elem, flatLvlNumbParam, tzData.FlatLvlNumbParamName);
        //                CheckMissingParam(elem, gripParam1, tzData.GripParamName1);
        //                if (gripParam2 != null)
        //                    CheckMissingParam(elem, gripParam2, tzData.GripParamName1);
        //            }

        //            // Забираю основные параметры
        //            Parameter flatNameParam = elem.LookupParameter(tzData.FlatNameParamName);
        //            Parameter areaParam = elem.get_Parameter(BuiltInParameter.ROOM_AREA);
        //            paramsToCheck.Add(areaParam);
        //            Parameter areaCoeffParam = elem.LookupParameter(tzData.AreaCoeffParamName);
        //            paramsToCheck.Add(areaCoeffParam);
        //            Parameter sumAreaCoeffParam = elem.LookupParameter(tzData.SumAreaCoeffParamName);
        //            paramsToCheck.Add(sumAreaCoeffParam);
        //            Parameter tzCodeNameParam = elem.LookupParameter(tzData.TZCodeParamName);
        //            paramsToCheck.Add(tzCodeNameParam);
        //            Parameter tzRangeNameParam = elem.LookupParameter(tzData.TZRangeNameParamName);
        //            paramsToCheck.Add(tzRangeNameParam);
        //            Parameter tzPercentParam = elem.LookupParameter(tzData.TZPercentParamName);
        //            paramsToCheck.Add(tzPercentParam);
        //            Parameter tzAreaMinParam = elem.LookupParameter(tzData.TZAreaMinParamName);
        //            paramsToCheck.Add(tzAreaMinParam);
        //            Parameter tzAreaMaxParam = elem.LookupParameter(tzData.TZAreaMaxParamName);
        //            paramsToCheck.Add(tzAreaMaxParam);
        //            Parameter modelPercentParamName = elem.LookupParameter(tzData.ModelPercentParamName);
        //            paramsToCheck.Add(modelPercentParamName);
        //            Parameter modelPercentToleranceParamName = elem.LookupParameter(tzData.ModelPercentToleranceParamName);
        //            paramsToCheck.Add(modelPercentToleranceParamName);


        //            // Проверка на пропущенные параметры
        //            CheckMissingParam(elem, flatNameParam, tzData.FlatNameParamName);
        //            CheckMissingParam(elem, areaParam, elem.get_Parameter(BuiltInParameter.ROOM_AREA).Definition.Name);
        //            CheckMissingParam(elem, areaCoeffParam, tzData.AreaCoeffParamName);
        //            CheckMissingParam(elem, sumAreaCoeffParam, tzData.SumAreaCoeffParamName);
        //            CheckMissingParam(elem, tzCodeNameParam, tzData.TZCodeParamName);
        //            CheckMissingParam(elem, tzRangeNameParam, tzData.TZRangeNameParamName);
        //            CheckMissingParam(elem, tzPercentParam, tzData.TZPercentParamName);
        //            CheckMissingParam(elem, tzAreaMinParam, tzData.TZAreaMinParamName);
        //            CheckMissingParam(elem, tzAreaMaxParam, tzData.TZAreaMaxParamName);
        //            CheckMissingParam(elem, modelPercentParamName, tzData.ModelPercentParamName);
        //            CheckMissingParam(elem, modelPercentToleranceParamName, tzData.ModelPercentToleranceParamName);

        //            // Общая сумма по проверяемым параметрам
        //            if (HasMissing(paramsToCheck))
        //                continue;

        //            // Получаю значения
        //            string flatNameData = GetParamValue(flatNameParam);
        //            double areaData = areaParam.AsDouble();
        //            string balkTerLogNameData = GetParamValue(balkTerLogNameParam);
        //            string flatNumbData = GetParamValue(flatNumbParam);
        //            string flatLvlNumbData = GetParamValue(flatLvlNumbParam);
        //            string gripData1 = GetParamValue(gripParam1);
        //            string gripData2 = GetParamValue(gripParam2);
        //            string tzCodeNameData = GetParamValue(tzCodeNameParam);

        //            // Проверка пустых значений
        //            CheckEmptyValue(elem, tzData.FlatNameParamName, flatNameData);
        //            if (!presetFlatCode)
        //                CheckEmptyValue(elem, tzData.TZCodeParamName, tzCodeNameData);

        //            if (tzData.HeatingRoomsInPrj)
        //            {
        //                CheckEmptyValue(elem, tzData.BalkTerLogNameParamName, balkTerLogNameData);
        //                CheckEmptyValue(elem, tzData.FlatNumbParamName, flatNumbData);
        //                CheckEmptyValue(elem, tzData.FlatLvlNumbParamName, flatLvlNumbData);
        //                CheckEmptyValue(elem, tzData.GripParamName1, gripData1);
        //                if (!string.IsNullOrEmpty(tzData.GripParamName2))
        //                    CheckEmptyValue(elem, tzData.GripParamName2, gripData2);
        //            }

        //            if (ErrorDict_Room.Values.Any(v => v.Contains(elem.Id)))
        //                continue;

        //            ARPG_Room aRPG_Room = new ARPG_Room()
        //            {
        //                Elem_Room = elem,
        //                AreaData_Room = areaData,
        //                ARPG_RoomType = roomRoomType,
        //                FlatNameData_Room = flatNameData,
        //                BalkTerLogNameData_Room = balkTerLogNameData,
        //                FlatNumbData_Room = flatNumbData,
        //                FlatLvlNumbData_Room = flatLvlNumbData,
        //                GripData1_Room = gripData1,
        //                GripData2_Room = gripData2,
        //                AreaCoeffParam_Room = areaCoeffParam,
        //                SumAreaCoeffParam_Room = sumAreaCoeffParam,
        //                TZCodeParam_Room = tzCodeNameParam,
        //                TZRangeNameParam_Room = tzRangeNameParam,
        //                TZPercentParam_Room = tzPercentParam,
        //                TZAreaMinParam_Room = tzAreaMinParam,
        //                TZAreaMaxParam_Room = tzAreaMaxParam,
        //                ModelPercentParam_Room = modelPercentParamName,
        //                ModelPercentToleranceParam_Room = modelPercentToleranceParamName,
        //            };

        //            aRPG_Room.Set_AreaCoeffAndFlatArea(tzData);
                    
        //            result.Add(aRPG_Room);
        //        }
        //        else
        //        {
        //            RoomType roomRoomType = GetRoomType(tzData, elem, false);

        //            List<Parameter> paramsToCheck = new List<Parameter>();
        //            // Если в проекте есть лоджии/балконы/террасы - нужно делать биндинг по помещениям
        //            Parameter balkTerLogNameParam = null;
        //            Parameter flatNumbParam = null;
        //            Parameter flatLvlNumbParam = null;
        //            Parameter gripParam1 = null;
        //            Parameter gripParam2 = null;

        //            // Забираю доп. параметры 
        //            if (tzData.HeatingRoomsInPrj)
        //            {
        //                balkTerLogNameParam = elem.LookupParameter(tzData.BalkTerLogNameParamName);
        //                paramsToCheck.Add(balkTerLogNameParam);
        //                flatNumbParam = elem.LookupParameter(tzData.FlatNumbParamName);
        //                paramsToCheck.Add(flatNumbParam);
        //                flatLvlNumbParam = elem.LookupParameter(tzData.FlatLvlNumbParamName);
        //                paramsToCheck.Add(flatLvlNumbParam);
        //                gripParam1 = elem.LookupParameter(tzData.GripParamName1);
        //                paramsToCheck.Add(gripParam1);
        //                if (!string.IsNullOrEmpty(tzData.GripParamName2))
        //                {
        //                    gripParam2 = elem.LookupParameter(tzData.GripParamName2);
        //                    paramsToCheck.Add(gripParam2);
        //                }

        //                // Проверка на пропущенные параметры
        //                CheckMissingParam(elem, balkTerLogNameParam, tzData.BalkTerLogNameParamName);
        //                CheckMissingParam(elem, flatNumbParam, tzData.FlatNumbParamName);
        //                CheckMissingParam(elem, flatLvlNumbParam, tzData.FlatLvlNumbParamName);
        //                CheckMissingParam(elem, gripParam1, tzData.GripParamName1);
        //                if (gripParam2 != null)
        //                    CheckMissingParam(elem, gripParam2, tzData.GripParamName1);
        //            }

        //            // Забираю основные параметры
        //            Parameter areaParam = elem.get_Parameter(BuiltInParameter.ROOM_AREA);
        //            paramsToCheck.Add(areaParam);
        //            Parameter areaCoeffParam = elem.LookupParameter(tzData.AreaCoeffParamName);
        //            paramsToCheck.Add(areaCoeffParam);
        //            Parameter sumAreaCoeffParam = elem.LookupParameter(tzData.SumAreaCoeffParamName);
        //            paramsToCheck.Add(sumAreaCoeffParam);
        //            Parameter tzCodeNameParam = elem.LookupParameter(tzData.TZCodeParamName);
        //            paramsToCheck.Add(tzCodeNameParam);
        //            Parameter tzRangeNameParam = elem.LookupParameter(tzData.TZRangeNameParamName);
        //            paramsToCheck.Add(tzRangeNameParam);
        //            Parameter tzPercentParam = elem.LookupParameter(tzData.TZPercentParamName);
        //            paramsToCheck.Add(tzPercentParam);
        //            Parameter tzAreaMinParam = elem.LookupParameter(tzData.TZAreaMinParamName);
        //            paramsToCheck.Add(tzAreaMinParam);
        //            Parameter tzAreaMaxParam = elem.LookupParameter(tzData.TZAreaMaxParamName);
        //            paramsToCheck.Add(tzAreaMaxParam);
        //            Parameter modelPercentParamName = elem.LookupParameter(tzData.ModelPercentParamName);
        //            paramsToCheck.Add(modelPercentParamName);
        //            Parameter modelPercentToleranceParamName = elem.LookupParameter(tzData.ModelPercentToleranceParamName);
        //            paramsToCheck.Add(modelPercentToleranceParamName);


        //            // Проверка на пропущенные параметры
        //            CheckMissingParam(elem, areaParam, elem.get_Parameter(BuiltInParameter.ROOM_AREA).Definition.Name);
        //            CheckMissingParam(elem, areaCoeffParam, tzData.AreaCoeffParamName);
        //            CheckMissingParam(elem, sumAreaCoeffParam, tzData.SumAreaCoeffParamName);
        //            CheckMissingParam(elem, tzCodeNameParam, tzData.TZCodeParamName);
        //            CheckMissingParam(elem, tzRangeNameParam, tzData.TZRangeNameParamName);
        //            CheckMissingParam(elem, tzPercentParam, tzData.TZPercentParamName);
        //            CheckMissingParam(elem, tzAreaMinParam, tzData.TZAreaMinParamName);
        //            CheckMissingParam(elem, tzAreaMaxParam, tzData.TZAreaMaxParamName);
        //            CheckMissingParam(elem, modelPercentParamName, tzData.ModelPercentParamName);
        //            CheckMissingParam(elem, modelPercentToleranceParamName, tzData.ModelPercentToleranceParamName);

        //            // Общая сумма по проверяемым параметрам
        //            if (HasMissing(paramsToCheck))
        //                continue;

        //            // Получаю значения
        //            double areaData = areaParam.AsDouble();
        //            string balkTerLogNameData = GetParamValue(balkTerLogNameParam);
        //            string flatNumbData = GetParamValue(flatNumbParam);
        //            string flatLvlNumbData = GetParamValue(flatLvlNumbParam);
        //            string gripData1 = GetParamValue(gripParam1);
        //            string gripData2 = GetParamValue(gripParam2);
        //            string tzCodeNameData = GetParamValue(tzCodeNameParam);

        //            // Проверка пустых значений
        //            if (!presetFlatCode)
        //                CheckEmptyValue(elem, tzData.TZCodeParamName, tzCodeNameData);

        //            if (tzData.HeatingRoomsInPrj)
        //            {
        //                CheckEmptyValue(elem, tzData.BalkTerLogNameParamName, balkTerLogNameData);
        //                CheckEmptyValue(elem, tzData.FlatNumbParamName, flatNumbData);
        //                CheckEmptyValue(elem, tzData.FlatLvlNumbParamName, flatLvlNumbData);
        //                CheckEmptyValue(elem, tzData.GripParamName1, gripData1);
        //                if (!string.IsNullOrEmpty(tzData.GripParamName2))
        //                    CheckEmptyValue(elem, tzData.GripParamName2, gripData2);
        //            }

        //            if (ErrorDict_Room.Values.Any(v => v.Contains(elem.Id)))
        //                continue;

        //            ARPG_Room aRPG_Room = new ARPG_Room()
        //            {
        //                Elem_Room = elem,
        //                AreaData_Room = areaData,
        //                ARPG_RoomType = roomRoomType,
        //                BalkTerLogNameData_Room = balkTerLogNameData,
        //                FlatNumbData_Room = flatNumbData,
        //                FlatLvlNumbData_Room = flatLvlNumbData,
        //                GripData1_Room = gripData1,
        //                GripData2_Room = gripData2,
        //                AreaCoeffParam_Room = areaCoeffParam,
        //                SumAreaCoeffParam_Room = sumAreaCoeffParam,
        //                TZCodeParam_Room = tzCodeNameParam,
        //                TZRangeNameParam_Room = tzRangeNameParam,
        //                TZPercentParam_Room = tzPercentParam,
        //                TZAreaMinParam_Room = tzAreaMinParam,
        //                TZAreaMaxParam_Room = tzAreaMaxParam,
        //                ModelPercentParam_Room = modelPercentParamName,
        //                ModelPercentToleranceParam_Room = modelPercentToleranceParamName,
        //            };

        //            aRPG_Room.Set_AreaCoeffAndFlatArea(tzData);

        //            result.Add(aRPG_Room);
        //        }
        //    }

        //    return result.ToArray();
        //}

        private static void CheckMissingParam(Element elem, Parameter param, string paramName)
        {
            if (param == null)
                HtmlOutput.SetMsgDict_ByMsg($"Отсутсвует параметр {paramName}", elem.Id, ErrorDict_Room);
        }

        private static void CheckEmptyValue(Element elem, string paramName, string value)
        {
            if (string.IsNullOrEmpty(value))
                HtmlOutput.SetMsgDict_ByMsg($"Пустой параметр {paramName}", elem.Id, ErrorDict_Room);
        }

        private static bool HasMissing(List<Parameter> parameters)
        {
            foreach (Parameter p in parameters)
            {
                if (p == null)
                    return true;
            }
            return false;
        }

        private static string GetParamValue(Parameter p)
        {
            try
            {
                if (p == null) return null;

                string val = p.AsValueString();
                if (!string.IsNullOrEmpty(val)) return val;

                string str = p.AsString();
                if (!string.IsNullOrEmpty(str)) return str;

                return null;
            }
            catch
            {
                return null;
            }
        }
        #endregion

        /// <summary>
        /// Установить тип помещения 
        /// </summary>
        private static RoomType GetRoomType(ARPG_TZ_MainData tzData, Element elem)
        {
            bool IsMatch(string a, string b) => DamerauLevenshteinDistance(a, b) <= 2;

            Parameter flatParam = elem.LookupParameter(tzData.FlatNameParamName);
            if (flatParam != null)
            {
                string flatData = flatParam.AsValueString();
                if (IsMatch(flatParam.AsValueString(), _roomName))
                    return RoomType.Flat;
                else if (IsMatch(flatData, _balkName))
                    return RoomType.Balkon;
                else if (IsMatch(flatData, _logName))
                    return RoomType.Loggia;
                else if (IsMatch(flatData, _terName))
                    return RoomType.Terrace;
            }

            return RoomType.Other;
        }

        /// <summary>
        /// Установить коэффициент и площадь с учётом коэффициента
        /// </summary>
        private ARPG_Room Set_AreaCoeffAndFlatArea(ARPG_TZ_MainData tzData)
        {
            double coeff = 1;

            switch (ARPG_RoomType)
            {
                case RoomType.Balkon:
                    double.TryParse(tzData.BalkAreaCoeff, out coeff);
                    break;
                case RoomType.Loggia:
                    double.TryParse(tzData.LogAreaCoeff, out coeff);
                    break;
                case RoomType.Terrace:
                    double.TryParse(tzData.TerraceAreaCoeff, out coeff);
                    break;
                case RoomType.Flat:
                    double.TryParse(tzData.FlatAreaCoeff, out coeff);
                    break;
            }

            double.TryParse(tzData.FlatAreaTolerance, out double tolerance);

            AreaCoeff_Room = coeff;
            FlatAreaTolerance_Room = tolerance;
            AreaCoeffData_Room = AreaData_Room * coeff;

            return this;
        }

        /// <summary>
        ///  Расчет расстояния Дамерлоу-Левинштейна
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

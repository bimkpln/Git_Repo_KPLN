using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools.Forms.Models.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common
{
    /// <summary>
    /// Контейнер данных из модели по ОТДЕЛЬНЫМ помещениям
    /// </summary>
    public sealed class ARPG_Room
    {
        private readonly string _balkName = "Балкон";
        private readonly string _logName = "Лоджия";

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
        /// Имя квартиры
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
        public string GripData1_Room { get; private set; }

        /// <summary>
        /// Захватка 2 (напр: секция)
        /// </summary>
        public string GripData2_Room { get; private set; }

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
        /// Получить коллекцию помещений в зависимости от анализируемого типа
        /// </summary>
        internal static ARPG_Room[] Get_ARPG_Rooms(Document doc, ARPG_TZ_MainData tzData)
        {
            if (tzData.FlatType == nameof(FilledRegion))
                return GetFromFilledRegions(doc);
            else if (tzData.FlatType == nameof(Room))
                return GetFromRooms(doc, tzData);
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
        private static ARPG_Room[] GetFromFilledRegions(Document doc)
        {
            FilledRegion[] filledRegions = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegion))
                .WhereElementIsNotElementType()
                .OfType<FilledRegion>()
                .Where(fr =>
                {
                    Parameter typeParam = fr.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    return typeParam != null && typeParam.AsValueString() != null && typeParam.AsValueString().StartsWith("ТЭП_");
                })
                .ToArray();

            // TODO: Реализовать при готовности тестовой модели
            return new ARPG_Room[0];
        }
        #endregion

        #region Помещения
        private static ARPG_Room[] GetFromRooms(Document doc, ARPG_TZ_MainData tzData)
        {
            ErrorDict_Room = new Dictionary<string, List<ElementId>>();

            Room[] rooms = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType()
                .OfType<Room>()
                .Where(r => r.Area > 0)
                .ToArray();

            List<ARPG_Room> result = new List<ARPG_Room>();
            foreach (Room room in rooms)
            {
                Parameter areaParam = room.get_Parameter(BuiltInParameter.ROOM_AREA);
                Parameter flatNameParam = room.LookupParameter(tzData.FlatNameParamName);
                Parameter flatNumbParam = room.LookupParameter(tzData.FlatNumbParamName);
                Parameter flatLvlNumbParam = room.LookupParameter(tzData.FlatLvlNumbParamName);
                Parameter gripParam1 = room.LookupParameter(tzData.GripParamName1);
                Parameter gripParam2 = room.LookupParameter(tzData.GripParamName2);
                Parameter areaCoeffParam = room.LookupParameter(tzData.AreaCoeffParamName);
                Parameter sumAreaCoeffParam = room.LookupParameter(tzData.SumAreaCoeffParamName);
                Parameter tzCodeNameParam = room.LookupParameter(tzData.TZCodeParamName);
                Parameter tzRangeNameParam = room.LookupParameter(tzData.TZRangeNameParamName);
                Parameter tzPercentParam = room.LookupParameter(tzData.TZPercentParamName);
                Parameter tzAreaMinParam = room.LookupParameter(tzData.TZAreaMinParamName);
                Parameter tzAreaMaxParam = room.LookupParameter(tzData.TZAreaMaxParamName);
                Parameter modelPercentParamName = room.LookupParameter(tzData.ModelPercentParamName);
                Parameter modelPercentToleranceParamName = room.LookupParameter(tzData.ModelPercentToleranceParamName);


                // Проверка на пропущенные параметры
                CheckMissingParam(room, areaParam, room.get_Parameter(BuiltInParameter.ROOM_AREA).Definition.Name);
                CheckMissingParam(room, flatNameParam, tzData.FlatNameParamName);
                CheckMissingParam(room, flatNumbParam, tzData.FlatNumbParamName);
                CheckMissingParam(room, flatLvlNumbParam, tzData.FlatLvlNumbParamName);
                CheckMissingParam(room, gripParam1, tzData.GripParamName1);
                CheckMissingParam(room, areaCoeffParam, tzData.AreaCoeffParamName);
                CheckMissingParam(room, sumAreaCoeffParam, tzData.SumAreaCoeffParamName);
                CheckMissingParam(room, tzCodeNameParam, tzData.TZCodeParamName);
                CheckMissingParam(room, tzRangeNameParam, tzData.TZRangeNameParamName);
                CheckMissingParam(room, tzPercentParam, tzData.TZPercentParamName);
                CheckMissingParam(room, tzAreaMinParam, tzData.TZAreaMinParamName);
                CheckMissingParam(room, tzAreaMaxParam, tzData.TZAreaMaxParamName);
                CheckMissingParam(room, modelPercentParamName, tzData.ModelPercentParamName);
                CheckMissingParam(room, modelPercentToleranceParamName, tzData.ModelPercentToleranceParamName);

                List<Parameter> paramsToCheck = new List<Parameter>()
                { 
                    areaParam, flatNumbParam, flatLvlNumbParam, gripParam1, areaCoeffParam, sumAreaCoeffParam,
                    tzCodeNameParam, tzRangeNameParam, tzPercentParam, tzAreaMinParam, tzAreaMaxParam,
                    modelPercentParamName, modelPercentToleranceParamName,
                };
                if (!string.IsNullOrEmpty(tzData.GripParamName2))
                {
                    CheckMissingParam(room, gripParam2, tzData.GripParamName2);
                    paramsToCheck.Add(gripParam2);
                }

                if (HasMissing(paramsToCheck))
                    continue;


                // Получаю значения
                var tzCodeNameData = GetParamValue(tzCodeNameParam);
                double areaData = areaParam.AsDouble();
                string flatNameData = GetParamValue(flatNameParam);
                string flatNumbData = GetParamValue(flatNumbParam);
                string flatLvlNumbData = GetParamValue(flatLvlNumbParam);
                string gripData1 = GetParamValue(gripParam1);
                string gripData2 = GetParamValue(gripParam2);

                // Проверка пустых значений
                CheckEmptyValue(room, tzData.TZCodeParamName, tzCodeNameData);
                CheckEmptyValue(room, tzData.FlatNameParamName, flatNameData);
                CheckEmptyValue(room, tzData.FlatNumbParamName, flatNumbData);
                CheckEmptyValue(room, tzData.FlatLvlNumbParamName, flatLvlNumbData);
                CheckEmptyValue(room, tzData.GripParamName1, gripData1);
                if (!string.IsNullOrEmpty(tzData.GripParamName2))
                    CheckEmptyValue(room, tzData.GripParamName2, gripData2);

                if (ErrorDict_Room.Values.Any(v => v.Contains(room.Id)))
                    continue;

                ARPG_Room aRPG_Room = new ARPG_Room()
                {
                    Elem_Room = room,
                    AreaData_Room = areaData,
                    FlatNameData_Room = flatNameData,
                    FlatNumbData_Room = flatNumbData,
                    FlatLvlNumbData_Room = flatLvlNumbData,
                    GripData1_Room = gripData1,
                    GripData2_Room = gripData2,
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

            return result.ToArray();
        }

        private static void CheckMissingParam(Room room, Parameter param, string paramName)
        {
            if (param == null)
                HtmlOutput.SetMsgDict_ByMsg($"Отсутсвует параметр {paramName}", room.Id, ErrorDict_Room);
        }

        private static void CheckEmptyValue(Room room, string paramName, string value)
        {
            if (string.IsNullOrEmpty(value))
                HtmlOutput.SetMsgDict_ByMsg($"Пустой параметр {paramName}", room.Id, ErrorDict_Room);
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
        /// Установить коэффициент и площадь с учётом коэффициента
        /// </summary>
        private ARPG_Room Set_AreaCoeffAndFlatArea(ARPG_TZ_MainData tzData)
        {
            double coeff;
            if (FlatNameData_Room.StartsWith(_balkName))
                double.TryParse(tzData.BalkAreaCoeff, out coeff);
            else if (FlatNameData_Room.StartsWith(_logName))
                double.TryParse(tzData.LogAreaCoeff, out coeff);
            else
                double.TryParse(tzData.FlatAreaCoeff, out coeff);

            AreaCoeff_Room = coeff;
            AreaCoeffData_Room = AreaData_Room * AreaCoeff_Room;

            return this;
        }
    }
}

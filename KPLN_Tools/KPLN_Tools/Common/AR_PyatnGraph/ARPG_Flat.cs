using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools.Forms.Models.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.AR_PyatnGraph
{
    /// <summary>
    /// Контейнер данных из модели по КВАРТИРАМ
    /// </summary>
    public sealed class ARPG_Flat
    {
        /// <summary>
        /// Словарь ОШИБОК элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        public static Dictionary<string, List<ElementId>> ErrorDict_Flat;

        private ARPG_Flat() { }

        /// <summary>
        /// Коллеция помещений в квартире
        /// </summary>
        public ARPG_Room[] ARPG_Rooms_Flat { get; private set; }

        /// <summary>
        /// Номер квартиры
        /// </summary>
        public string FlatNumbData_Flat { get; private set; }

        /// <summary>
        /// Номер этажа квартиры
        /// </summary>
        public string FlatLvlNumbData_Flat { get; private set; }

        /// <summary>
        /// Захватка 1 (напр: корпус)
        /// </summary>
        public string GripData1_Flat { get; private set; }

        /// <summary>
        /// Захватка 2 (напр: секция)
        /// </summary>
        public string GripData2_Flat { get; private set; }

        /// <summary>
        /// Получить коллекцию квартир
        /// </summary>
        internal static ARPG_Flat[] Get_ARPG_Flats(ARPG_TZ_MainData tzData, ARPG_Room[] arpgRooms)
        {
            if (arpgRooms == null || arpgRooms.Length == 0)
                return new ARPG_Flat[0];

            if (tzData.HeatingRoomsInPrj)
            {
                var flats = arpgRooms
                    .OrderBy(r => r.GripData1_Room)
                    .ThenBy(r => r.GripData2_Room)
                    .ThenBy(r => r.FlatLvlNumbData_Room)
                    .ThenBy(r => r.FlatNumbData_Room)
                    .GroupBy(r => new
                    {
                        r.GripData1_Room,
                        r.GripData2_Room,
                        r.FlatLvlNumbData_Room,
                        r.FlatNumbData_Room
                    })
                    .Select(g => new ARPG_Flat
                    {
                        GripData1_Flat = g.Key.GripData1_Room,
                        GripData2_Flat = g.Key.GripData2_Room,
                        FlatLvlNumbData_Flat = g.Key.FlatLvlNumbData_Room,
                        FlatNumbData_Flat = g.Key.FlatNumbData_Room,
                        ARPG_Rooms_Flat = g.ToArray()
                    })
                    .ToArray();

                return flats;
            }
            else
            {
                return arpgRooms
                    .Select(ar => new ARPG_Flat { ARPG_Rooms_Flat = new ARPG_Room[] { ar } })
                    .ToArray();
            }
        }

        /// <summary>
        /// Задать основные первичные параметры
        /// </summary>
        internal static void SetMainFlatData(ARPG_TZ_MainData arpgTZMainData, ARPG_Flat[] arpgFlats, ARPG_TZ_FlatData[] arpgTZFlatDatas)
        {
            ErrorDict_Flat = new Dictionary<string, List<ElementId>>();

            foreach (ARPG_Flat arpgFlat in arpgFlats)
            {
                ARPG_Room[] arpgRooms = arpgFlat.ARPG_Rooms_Flat;

                // Рассчёт площади с коэф
                double flatSumArea = arpgRooms.Sum(x => x.AreaCoeffData_Room);
#if Debug2020 || Revit2020
                double flatSumArea_SqM = UnitUtils.ConvertFromInternalUnits(flatSumArea, DisplayUnitType.DUT_SQUARE_METERS);
#else
                double flatSumArea_SqM = UnitUtils.ConvertFromInternalUnits(flatSumArea, UnitTypeId.SquareMeters);
#endif
                //// Получаю квартиру по условиям ТЗ по площади
                //ARPG_TZ_FlatData arpgTZFlatDataByTZ = arpgTZFlatDatas.FirstOrDefault(tz => tz.TZAreaMin_Double <= flatSumArea_SqM && tz.TZAreaMax_Double >= flatSumArea_SqM);
                //if (arpgTZFlatDataByTZ == null)
                //{
                //    foreach (ARPG_Room arpgRoom in arpgRooms)
                //    {
                //        HtmlOutput.SetMsgDict_ByMsg(
                //            $"Помещения из квартиры {arpgFlat.FlatNumbData_Flat} " +
                //                $"на этаже {arpgFlat.FlatLvlNumbData_Flat} по захватке {arpgFlat.GripData1_Flat} " +
                //                $"ВНЕ диапазонов из ТЗ. Сейчас суммарная площадь: {Math.Round(flatSumArea_SqM, 3)} м2",
                //            arpgRoom.Elem_Room.Id, ErrorDict_Flat);
                //    }
                //}

                // Проверяю и получаю квартиру по условиям ТЗ по коду (если нужно)
                ARPG_TZ_FlatData arpgTZFlatDataByTZ = null;
                if (arpgTZMainData.HeatingRoomsInPrj)
                {
                    HashSet<string> arpgRoomTZCodeDataColl = arpgRooms.Select(ar => ar.TZCodeParam_Room.AsString()).ToHashSet();
                    if (arpgRoomTZCodeDataColl.Count != 1)
                    {
                        foreach (ARPG_Room arpgRoom in arpgRooms)
                        {
                            HtmlOutput.SetMsgDict_ByMsg(
                                $"Помещения из квартиры {arpgFlat.FlatNumbData_Flat} " +
                                    $"на этаже {arpgFlat.FlatLvlNumbData_Flat} по захватке {arpgFlat.GripData1_Flat} " +
                                    $"попали в одну квартиру, но при этом - должны быть в разных квартирах (см. параметр кода квартиры).",
                                arpgRoom.Elem_Room.Id, ErrorDict_Flat);
                        }

                        continue;
                    }

                    // Проверяю код квартиры с кодом из ТЗ
                    arpgTZFlatDataByTZ = arpgTZFlatDatas.FirstOrDefault(tz => tz.TZCode == arpgRoomTZCodeDataColl.FirstOrDefault());
                    if (arpgTZFlatDataByTZ == null)
                    {
                        foreach (ARPG_Room arpgRoom in arpgRooms)
                        {
                            HtmlOutput.SetMsgDict_ByMsg(
                                $"Помещения из квартиры {arpgFlat.FlatNumbData_Flat} " +
                                    $"на этаже {arpgFlat.FlatLvlNumbData_Flat} по захватке {arpgFlat.GripData1_Flat} " +
                                    $"НЕ соответсвуют принятому коду из ТЗ, сейчас заполнено \"{arpgRoomTZCodeDataColl.FirstOrDefault()}\"",
                                arpgRoom.Elem_Room.Id, ErrorDict_Flat);
                        }
                    
                        continue;
                    }
                }
                else
                    arpgTZFlatDataByTZ = arpgTZFlatDatas.FirstOrDefault(tz => tz.TZCode == arpgRooms[0].TZCodeParam_Room.AsString());


                // Проверяю на совпадение диапазона квартиры с диапазоном ТЗ и с принятым допуском (допуск у всех одинаковый)
                double areaTolerance = arpgRooms.FirstOrDefault().FlatAreaTolerance_Room;
                ARPG_TZ_FlatData arpgTZFlatDataBySumArea = arpgTZFlatDatas.FirstOrDefault(tz => 
                    tz.TZAreaMin_Double - areaTolerance <= flatSumArea_SqM 
                    && tz.TZAreaMax_Double + areaTolerance >= flatSumArea_SqM);
                
                if (arpgTZFlatDataBySumArea == null || arpgTZFlatDataByTZ.TZCode != arpgTZFlatDataBySumArea.TZCode)
                {
                    foreach (ARPG_Room arpgRoom in arpgRooms)
                    {
                        HtmlOutput.SetMsgDict_ByMsg(
                            $"Помещения из квартиры {arpgFlat.FlatNumbData_Flat} " +
                                $"на этаже {arpgFlat.FlatLvlNumbData_Flat} по захватке {arpgFlat.GripData1_Flat} " +
                                $"ВНЕ диапазонов из ТЗ (напомню - допуск составляет \"{areaTolerance}\" м2." +
                                $"Сейчас суммарная площадь: {Math.Round(flatSumArea_SqM, 3)} м2, выбранный тип квартиры: {arpgTZFlatDataByTZ.TZCode}",
                            arpgRoom.Elem_Room.Id, ErrorDict_Flat);
                    }

                    continue;
                }


                // Если есть в списке замечаний - пропускаем
                if (ErrorDict_Flat.Values.SelectMany(v => v).Intersect(arpgRooms.Select(ar => ar.Elem_Room.Id)).Any())
                    continue;


                // Заполняю значения
                foreach (ARPG_Room arpgRoom in arpgRooms)
                {
                    // Данные из ТЗ 
                    arpgRoom.TZCodeParam_Room.Set(arpgTZFlatDataBySumArea.TZCode);
                    arpgRoom.TZRangeNameParam_Room.Set(arpgTZFlatDataBySumArea.TZRangeName);


                    if (double.TryParse(arpgTZFlatDataBySumArea.TZPercent, out double tzPercent))
                        arpgRoom.TZPercentParam_Room.Set(tzPercent);
                    else
                        throw new Exception($"Отправь разработчику - в анализ попали данные, которые нельзя преобразовать в double: {arpgTZFlatDataByTZ.TZPercent}");

#if Debug2020 || Revit2020
                    arpgRoom.TZAreaMinParam_Room.Set(UnitUtils.ConvertToInternalUnits(arpgTZFlatDataBySumArea.TZAreaMin_Double, DisplayUnitType.DUT_SQUARE_METERS));
                    arpgRoom.TZAreaMaxParam_Room.Set(UnitUtils.ConvertToInternalUnits(arpgTZFlatDataBySumArea.TZAreaMax_Double, DisplayUnitType.DUT_SQUARE_METERS));
#else
                    arpgRoom.TZAreaMinParam_Room.Set(UnitUtils.ConvertToInternalUnits(arpgTZFlatDataBySumArea.TZAreaMin_Double, UnitTypeId.SquareMeters));
                    arpgRoom.TZAreaMaxParam_Room.Set(UnitUtils.ConvertToInternalUnits(arpgTZFlatDataBySumArea.TZAreaMax_Double, UnitTypeId.SquareMeters));
#endif

                    // Площади
                    arpgRoom.AreaCoeffParam_Room.Set(arpgRoom.AreaCoeffData_Room);
                    arpgRoom.SumAreaCoeffParam_Room.Set(flatSumArea);
                }
            }
        }
    }
}

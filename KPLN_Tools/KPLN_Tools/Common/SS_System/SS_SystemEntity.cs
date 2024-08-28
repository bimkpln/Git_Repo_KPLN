using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.SS_System
{
    internal class SS_SystemEntity
    {
        /// <summary>
        /// Имя параметра: 'КП_И_Адрес текущий'
        /// </summary>
        public static readonly string CurrentAdressParamName = "КП_И_Адрес текущий";
        /// <summary>
        /// Имя параметра: 'КП_И_Адрес предыдущий'
        /// </summary>
        public static readonly string PreviousAdressParamName = "КП_И_Адрес предыдущий";
        /// <summary>
        /// Имя параметра: 'КП_И_Количество занимаемых адресов'
        /// </summary>
        public static readonly string AdressCountParamName = "КП_И_Количество занимаемых адресов";
        /// <summary>
        /// Имя параметра: 'КП_О_Позиция'
        /// </summary>
        public static readonly string PositionParamName = "КП_О_Позиция";
        /// <summary>
        /// Имя параметра: 'КП_И_Префикс системы'
        /// </summary>
        public static readonly string SysPrefixParamName = "КП_И_Префикс системы";
        /// <summary>
        /// Имя параметра: 'КП_И_Длина линии'
        /// </summary>
        public static readonly string LenghtParamName = "КП_И_Длина линии";

        /// <summary>
        /// Коллекция используемых параметров
        /// </summary>
        public static readonly string[] EntityParams = new string[]
        {
            CurrentAdressParamName,
            PreviousAdressParamName,
            AdressCountParamName,
            PositionParamName,
            SysPrefixParamName,
            LenghtParamName,
        };

        public SS_SystemEntity(FamilyInstance famInst, ElectricalSystemType sysType)
        {
            CurrentFamInst = famInst;
            #if Revit2023 || Debug
            CurrentElSystemSet = CurrentFamInst
                .MEPModel
                .GetElectricalSystems()
                .Where(s => s.SystemType.Equals(sysType));
            #endif
        }

        /// <summary>
        /// Ссылка на FamilyInstance
        /// </summary>
        public FamilyInstance CurrentFamInst { get; private set; }

        /// <summary>
        /// Коллекция ElectricalSystem у элемента (из-за принятого метода моделирования - их в среднем 2 шт.)
        /// </summary>
        public IEnumerable<ElectricalSystem> CurrentElSystemSet { get; set; }

        /// <summary>
        /// Текущий адрес эл-та
        /// </summary>
        public string CurrentAdressData { get; private set; }

        /// <summary>
        /// Индекс системы
        /// </summary>
        public string CurrentAdressSystemIndex { get; private set; }

        /// <summary>
        /// Разделитель, используемый в адресации
        /// </summary>
        public string CurrentAdressSeparator { get; private set; }

        /// <summary>
        /// Индекс/порядковый номер элемента в цепи СС
        /// </summary>
        public int CurrentAdressIndex{ get; private set; }

        /// <summary>
        /// Количество занимаемых адресов
        /// </summary>
        public int CurrentAdressCountData { get; private set; }

        /// <summary>
        /// Получить данные параметров по элементу
        /// </summary>
        /// <exception cref="Exception"></exception>
        public SS_SystemEntity SetSysParamsInstData()
        {
            Parameter currentSysNameParam = CurrentFamInst.LookupParameter(CurrentAdressParamName) ??
                throw new Exception($"У элемента с id: {CurrentFamInst.Id} нет параметра {CurrentAdressParamName}");
            CurrentAdressData = currentSysNameParam.AsString();

            string[] splitedNumber = NumberService.SystemNumberSplit(CurrentAdressData);
            CurrentAdressIndex = int.Parse(splitedNumber[0]);
            CurrentAdressSystemIndex = splitedNumber[1];
            CurrentAdressSeparator = splitedNumber[2];

            Parameter currentAdressCountParam = CurrentFamInst.LookupParameter(AdressCountParamName) ??
                throw new Exception($"У элемента с id: {CurrentFamInst.Id} нет параметра {AdressCountParamName}");
            double adrDData = currentAdressCountParam.AsDouble();
            int adrIData = Convert.ToInt32(adrDData);
            CurrentAdressCountData = adrIData == 0 ? 1 : adrIData;

            return this;
        }
    }
}

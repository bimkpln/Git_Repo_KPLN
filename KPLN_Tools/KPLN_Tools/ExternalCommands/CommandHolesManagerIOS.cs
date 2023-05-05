using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.HolesManager;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandHolesManagerIOS : IExternalCommand
    {
        private Stopwatch _stopWatchUnBook = new Stopwatch();
        private readonly string _offsetUpParmaName = "SYS_OFFSET_UP";
        private readonly string _offsetDownParmaName = "SYS_OFFSET_DOWN";
        private readonly Guid _relativeElevParam = new Guid("76ce59d6-ad59-4684-a037-267930dd937d");
        private readonly Guid _absoluteElevParam = new Guid("7aae06b7-50b2-499c-81e6-aff0fd46b68b");
        private readonly Guid _familyVersionParam = new Guid("85cd0032-c9ee-4cd3-8ffa-b2f1a05328e3");
        private readonly Guid _revalueOffsetsParam = new Guid("15187bb1-249b-4ff8-9601-7124fba37ca7");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Получаю связанные модели АР (нужно доработать, т.к. сейчас возможны ошибки поиска моделей - лучше добавить проверку по БД) и элементы стяжек пола
                IEnumerable<RevitLinkInstance> linkedModels = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Where(lm => lm.Name.Contains("_AR_") || lm.Name.Contains("_АР_"))
                    .Cast<RevitLinkInstance>();
                if (!linkedModels.Any())
                    throw new Exception("Для работы обязателено нужно подгрузить все связи АР (кроме разбивочных файлов), которые являются подложкой для модели ИОС");

                // Получаю абсолютную отметку базовой точки проекта (т.е. абс. отметка проекта)
                var basePoint = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                    .WhereElementIsNotElementType()
                    .Cast<BasePoint>()
                    .FirstOrDefault();
                double absoluteElevBasePnt = basePoint.SharedPosition.Z;

                #region  Получаю и проверяю коллекцию классов-отверстий в стенах
                IEnumerable<FamilyInstance> holesElems = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .Cast<FamilyInstance>()
                    .Where(e =>
                        e.Symbol.FamilyName.StartsWith("501_ЗИ_Отверстие")
                        || e.Symbol.FamilyName.StartsWith("501_MEP_")
                        && !(e.Symbol.FamilyName.Contains("Перекрытие"))
                    );

                CheckMainParamsError(holesElems, 2, 0);
                CheckHolesExpandParamsError(holesElems);
                #endregion

                #region  Получаю и проверяю коллекцию классов-шахт/отверстий в перекрытиях
                IEnumerable<FamilyInstance> shaftElems = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .Cast<FamilyInstance>()
                    .Where(e =>
                    e.Symbol.FamilyName.StartsWith("501_ЗИ_Шахта")
                    || (e.Symbol.FamilyName.StartsWith("501_") && e.Symbol.FamilyName.Contains("Перекрытие"))
                );

                CheckMainParamsError(shaftElems, 1, 1);
                #endregion

                #region Подготовка и обработка спец. классов для последующей записи в проект
                IOSHolesPrepareManager iOSHolesPrepareManager = new IOSHolesPrepareManager(holesElems, linkedModels);
                List<IOSHoleDTO> holesDTOColl = iOSHolesPrepareManager.PrepareHolesDTO();

                IOSShaftPrepareManager iOSShaftPrepareManager = new IOSShaftPrepareManager(shaftElems, linkedModels);
                List<IOSShaftDTO> shaftDTOElems = iOSShaftPrepareManager.PrepareShaftDTO();
                #endregion

                #region Вывод ошибок пользователю
                if (iOSHolesPrepareManager.ErrorFamInstColl.Any())
                {
                    Print("Ошибки определения основы у отверстий ↑", KPLN_Loader.Preferences.MessageType.Header);
                    foreach (FamilyInstance fi in iOSHolesPrepareManager.ErrorFamInstColl)
                    {
                        Print($"KPLN: Ошибка - отверстие с id: {fi.Id} невозможно определить основу (уровень, на котором оно расположено). " +
                            $"Отверстие должно быть в стене, а стена должна стоять на Жб перекрытии, расстояние до которой не больше 15 м", KPLN_Loader.Preferences.MessageType.Warning);
                    }
                }

                if (iOSShaftPrepareManager.ErrorFamInstColl.Any())
                {
                    Print("Ошибки определения основы у шахт ↑", KPLN_Loader.Preferences.MessageType.Header);
                    foreach (FamilyInstance fi in iOSShaftPrepareManager.ErrorFamInstColl)
                    {
                        Print($"KPLN: Ошибка - шахта с id: {fi.Id} невозможно определить основу (уровень, на котором оно расположено). " +
                            $"Основание шахты должно быть либо в перекрытии, либо над ним не выше, чем на 0.300 м", KPLN_Loader.Preferences.MessageType.Warning);
                    }
                }
                #endregion

                #region Запись полученных данных в проект
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("КП_Отверстия для АР: границы и отметки");

                    HolesParamsWriter(holesDTOColl);
                    ShaftParamsWriter(shaftDTOElems);

                    t.Commit();
                }
                #endregion
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.InnerException.Message}\n StackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Header);
                else
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.Message}\n StackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Header);

                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Проверка наличия основных общих параметров семейств
        /// </summary>
        /// <returns></returns>
        private void CheckMainParamsError(IEnumerable<FamilyInstance> collElems, int fiMajorVers, int fiMinorVers)
        {
            if (collElems.Count() == 0)
                return;

            //// Проверка на наличие и значение параметра КП_О_Версия семейства
            //Dictionary<Family, Parameter> fsParamDict = new Dictionary<Family, Parameter>(collElems
            //    .ToDictionary(fi => fi.Symbol.Family, fi => fi.Symbol.get_Parameter(_familyVersionParam)));

            //foreach (KeyValuePair<Family, Parameter> kvp in fsParamDict)
            //{
            //    if (kvp.Value is null)
            //        throw new Exception($"KPLN: Ошибка - семейство {kvp.Key.Name} не актуальное. Обнови его отсюда {@"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM"}");
            //    else
            //    {
            //        string strVers = kvp.Value.AsString();
            //        string majorStrVersion = strVers[0].ToString();
            //        string minorStrVersion = strVers[strVers.Length - 1].ToString();
            //        if (int.TryParse(majorStrVersion, out int majorVersion) && int.TryParse(minorStrVersion, out int minorVersion))
            //        {
            //            if (majorVersion != fiMajorVers || !(minorVersion > fiMinorVers))
            //                throw new Exception($"KPLN: Ошибка - семейство {kvp.Key.Name} не актуальное. Обнови его отсюда {@"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\BIM"}");
            //        }
            //    }
            //}

            // Проверка на наличие параметра 00_Отметка_Абсолютная
            HashSet<Family> nullParamsFamilies = new HashSet<Family>(collElems
                .Where(fi => fi.get_Parameter(_relativeElevParam) is null || fi.get_Parameter(_absoluteElevParam) is null)
                .Select(fi => fi.Symbol.Family));
            if (nullParamsFamilies.Any())
            {
                foreach (Family f in nullParamsFamilies)
                    throw new Exception($"В семействе {f.Name} нет параметров экземпляра для записи абсолютной и относительной отметок");
            }
        }

        /// <summary>
        /// Проверка параметров оффсетов для овтерстий в стенах
        /// </summary>
        /// <param name="collElems"></param>
        /// <exception cref="Exception"></exception>
        private void CheckHolesExpandParamsError(IEnumerable<FamilyInstance> collElems)
        {
            if (collElems.Count() == 0)
                return;

            // Проверка на наличие параметра SYS_OFFSET_UP и SYS_OFFSET_DOWN
            HashSet<Family> nullParamsFamilies = new HashSet<Family>(collElems
                .Where(fi =>
                    (fi.LookupParameter(_offsetUpParmaName) is null || fi.LookupParameter(_offsetDownParmaName) is null))
                .Select(fi => fi.Symbol.Family));
            if (nullParamsFamilies.Any())
            {
                foreach (Family f in nullParamsFamilies)
                    throw new Exception($"В семействе {f.Name} нет параметра экземпляра {_offsetUpParmaName}, или {_offsetDownParmaName}");
            }
        }

        /// <summary>
        /// Задает значения для расширения границ и записывает значения парамтеров для отверстий из HoleDTO
        /// </summary>
        /// <param name="holesElems">Коллекция отверстий</param>
        private void HolesParamsWriter(IEnumerable<IOSHoleDTO> holesElems)
        {
            foreach (IOSHoleDTO dto in holesElems)
            {
                // Оффсеты
                var revalueParam = dto.CurrentHole.get_Parameter(_revalueOffsetsParam);
                if (revalueParam == null || revalueParam.AsInteger() == 1)
                {
                    dto.CurrentHole
                        .LookupParameter(_offsetDownParmaName)
                        .Set(dto.DownFloorDistance);

                    dto.CurrentHole
                        .LookupParameter(_offsetUpParmaName)
                        .Set(dto.UpFloorDistance);
                }

                // Отметки
                double rlvElev = RoundElevationData(dto.RlvElevation);
                double rlvHostElev = RoundElevationData(dto.DownBindingElevation);
                double absElev = RoundElevationData(dto.AbsElevation);

                string rlvElevStr = PrepareStringData(rlvElev);
                string rlvHostElevStr = PrepareStringData(rlvHostElev);
                string absElevStr = PrepareStringData(absElev);

                SetRlvElevation(dto.CurrentHole, dto.BindingPrefixString, rlvElevStr, rlvHostElevStr);
                SetAbsElevation(dto.CurrentHole, dto.BindingPrefixString, absElevStr);
            }
        }

        /// <summary>
        /// Запись значения парамтеров для шахт
        /// </summary>
        /// <param name="holesElems">Коллекция шахт</param>
        private void ShaftParamsWriter(IEnumerable<IOSShaftDTO> shaftElems)
        {
            foreach (IOSShaftDTO dto in shaftElems)
            {
                // Отметки
                double rlvElev = RoundElevationData(dto.RlvElevation);
                double rlvHostElev = RoundElevationData(dto.DownBindingElevation);
                double absElev = RoundElevationData(dto.AbsElevation);

                string rlvElevStr = PrepareStringData(rlvElev);
                string rlvHostElevStr = PrepareStringData(rlvHostElev);
                string absElevStr = PrepareStringData(absElev);

                SetRlvElevation(dto.CurrentHole, dto.BindingPrefixString, rlvElevStr, rlvHostElevStr);
                SetAbsElevation(dto.CurrentHole, dto.BindingPrefixString, absElevStr);
            }
        }

        private double RoundElevationData(double elev) => (Math.Round((elev * 304.8 / 5), 0) * 5) / 1000;

        private string PrepareStringData(double elev) => elev >= 0 ? string.Format("+{0:F3}", elev) : string.Format("{0:F3}", elev);

        private void SetRlvElevation(FamilyInstance fi, string prefix, string elemElev, string hostElev) =>
            fi.get_Parameter(_relativeElevParam).Set($"[Отн]: {prefix} {elemElev} отн. {hostElev}");

        private void SetAbsElevation(FamilyInstance fi, string prefix, string elemElev) =>
            fi.get_Parameter(_absoluteElevParam).Set($"[Абс]: {prefix} {elemElev} отн. нуля здания");
    }
}

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.HolesManager;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
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
                // Получаю абсолютную отметку базовой точки проекта (т.е. абс. отметка проекта)
                var basePoint = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                    .WhereElementIsNotElementType()
                    .Cast<BasePoint>()
                    .FirstOrDefault();
                double absoluteElevBasePnt = basePoint.SharedPosition.Z;

                #region  Получаю и проверяю коллекцию семейств-отверстий в стенах
                IEnumerable<FamilyInstance> holesElems = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .Cast<FamilyInstance>()
                    .Where(e => e.Symbol.FamilyName.StartsWith("501_ЗИ_Отверстие") || e.Symbol.FamilyName.StartsWith("501_MEP_"));
                
                CheckMainParamsError(holesElems, 2, 0);
                CheckHolesExpandParamsError(holesElems);
                #endregion

                #region  Получаю и проверяю коллекцию семейств-шахт/отверстий в перекрытиях
                IEnumerable<FamilyInstance> shaftElemsColl = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .Cast<FamilyInstance>()
                    .Where(e => e.Symbol.FamilyName.StartsWith("501_ЗИ_Шахта"));
                
                CheckMainParamsError(shaftElemsColl, 1, 1);
                #endregion

                #region Параллельная подготовка спец. классов для последующей записи в проект
                List<HoleDTO> holesDTOColl = new List<HoleDTO>();
                Task taskHoles = Task.Run(() =>
                {
                    // Получаю связанные модели АР (нужно доработать, т.к. сейчас возможны ошибки поиска моделей - лучше добавить проверку по БД) и элементы стяжек пола
                    IEnumerable<RevitLinkInstance> linkedModels = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Where(lm => lm.Name.Contains("_AR_") || lm.Name.Contains("_АР_"))
                        .Cast<RevitLinkInstance>();
                    
                    holesDTOColl = HolesPrepareManager.PrepareHolesDTO(holesElems, linkedModels, absoluteElevBasePnt);
                });

                List<ShaftDTO> shaftDTOElems = new List<ShaftDTO>();
                Task taskShafts = Task.Run(() =>
                {
                    // Получаю связанные модели АР (нужно доработать, т.к. сейчас возможны ошибки поиска моделей - лучше добавить проверку по БД) и элементы стяжек пола
                    IEnumerable<RevitLinkInstance> linkedModels = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Where(lm => lm.Name.Contains("_AR_") || lm.Name.Contains("_АР_"))
                        .Cast<RevitLinkInstance>();
                    shaftDTOElems = ShaftPrepareManager.PrepareShaftDTO(shaftElemsColl, linkedModels, absoluteElevBasePnt);
                });

                taskHoles.Wait();
                taskShafts.Wait();
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
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.InnerException.Message}", KPLN_Loader.Preferences.MessageType.Header);
                else
                    Print($"Работа скрипта остановлена. Устрани ошибку:\n {ex.Message}", KPLN_Loader.Preferences.MessageType.Header);

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
                    throw new Exception($"KPLN: Ошибка - в семействе {f.Name} нет параметров экземпляра для записи абсолютной и относительной отметок");
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
                    throw new Exception($"KPLN: Ошибка - в семействе {f.Name} нет параметра экземпляра {_offsetUpParmaName}, или {_offsetDownParmaName}");
            }
        }

        /// <summary>
        /// Задает значения для расширения границ и записывает значения парамтеров для отверстий из HoleDTO
        /// </summary>
        /// <param name="holesElems">Коллекция отверстий</param>
        private void HolesParamsWriter(IEnumerable<HoleDTO> holesElems)
        {
            foreach (HoleDTO dto in holesElems)
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
                dto.CurrentHole
                    .get_Parameter(_relativeElevParam)
                    .Set($"[Отн]: {dto.BindingPrefixString} {Math.Round((dto.RlvElevation * 304.8 / 5), 0) * 5} отн. {Math.Round((dto.DownBindingElevation * 304.8 / 5), 0) * 5}");
                
                dto.CurrentHole
                    .get_Parameter(_absoluteElevParam)
                    .Set($"[Абс]: {dto.BindingPrefixString} {Math.Round((dto.AbsElevation * 304.8 / 5), 0) * 5}");
            }
        }

        /// <summary>
        /// Запись значения парамтеров для шахт
        /// </summary>
        /// <param name="holesElems">Коллекция шахт</param>
        private void ShaftParamsWriter(IEnumerable<ShaftDTO> shaftElems)
        {
            foreach (ShaftDTO dto in shaftElems)
            {
                // Отметки
                dto.CurrentHole
                    .get_Parameter(_relativeElevParam)
                    .Set($"[Отн]: {dto.BindingPrefixString} {Math.Round((dto.RlvElevation * 304.8 / 5), 0) * 5} отн. {Math.Round((dto.DownBindingElevation * 304.8 / 5), 0) * 5}");

                dto.CurrentHole
                    .get_Parameter(_absoluteElevParam)
                    .Set($"[Абс]: {dto.BindingPrefixString} {Math.Round((dto.AbsElevation * 304.8 / 5), 0) * 5}");

            }
        }
    }
}

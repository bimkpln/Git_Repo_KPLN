using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckFamilies : AbstrCheckCommandOld<CommandCheckFamilies>, IExternalCommand
    {
        internal const string PluginName = "Проверка семейств";

        private readonly string[] _systemTypeCode = new string[]
        {
            // ЖЕЛЕЗОБЕТОН, КЛАДКА (АР/КР):
            "ЖБ",
            "К",
            "ККЛ",
            "ГБ",
            "ГСБ",
            "КБ",
            "СКЦ",
            "ПГП",
            "ПГПВ",
            "КАМ",
            // ИЗОЛЯЦИЯ (АР/КР):
            "ППС",
            "МИН",
            "ГИ",
            "ПИ",
            "ДРЕН",
            "ГЕО",
            "В",
            // ОТДЕЛКА - СТЕНЫ, ПОЛЫ, ПОТОЛКИ (АР/КР):
            "ГКЛ",
            "ГКЛВ",
            "ГВЛ",
            "ГВЛВ",
            "АЦЭИД",
            "ХЦЛ",
            "ЦСП",
            "ШТ",
            "ШТВ",
            "ШТД",
            "ШП",
            "ГР",
            "КРАС",
            "ОБ",
            "ФИБ",
            "ККЛ",
            "ПЛКЛ",
            "ПЛКР",
            "КРГ",
            "ПЛТР",
            "ПЛТАК",
            "ТЕРД  ",
            "КЛЕ",
            "ЦПС",
            "ЦПСА",
            "БС",
            "НП",
            "ПАР",
            "КВР",
            "ЛИН",
            "ЛАМ",
            "АРМС",
            "ГРИЛ",
            "НАТП",
            "АКУС",
            "КАМ",
        };

        public CommandCheckFamilies() : base()
        {
        }

        internal CommandCheckFamilies(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Application app = uiapp.Application;

            app.FailuresProcessing += FailuresProcessor;

            try
            {
                // Получаю коллекцию элементов для анализа
                FilteredElementCollector docFamsColl = new FilteredElementCollector(doc).OfClass(typeof(Family));
                FilteredElementCollector wallTypesColl = new FilteredElementCollector(doc).OfClass(typeof(WallType));
                FilteredElementCollector floorTypesColl = new FilteredElementCollector(doc).OfClass(typeof(FloorType));

                FilteredElementCollector resultColl = docFamsColl.UnionWith(wallTypesColl).UnionWith(floorTypesColl);

                #region Проверяю и обрабатываю элементы
                WPFEntity[] wpfColl = CheckCommandRunner(doc, resultColl.ToArray());
                OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
                if (form != null) form.Show();
                else return Result.Cancelled;
                #endregion

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Дополнительная обертка из дополнительного try/cath ради отписки от события FailuresProcessor
                throw ex;
            }
            finally
            {
                app.FailuresProcessing -= FailuresProcessor;
            }
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] elemColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (Element elem in elemColl)
            {
                // Проверяю семейства и их типоразмеры
                if (elem is Family currentFam)
                {
                    result.AddRange(CheckFamilyAndTypeDuplicateName(currentFam, elemColl));

                    WPFEntity checkFamilyPath = CheckFamilyPath(doc, currentFam);
                    if (checkFamilyPath != null)
                        result.Add(checkFamilyPath);
                }

                // Проверяю системные типоразмеры АР и КР
                else if (doc.PathName.Contains("АР_")
                    || doc.PathName.Contains("_АР_")
                    || doc.PathName.Contains("AR_")
                    || doc.PathName.Contains("_AR_")
                    || doc.PathName.Contains("КР_")
                    || doc.PathName.Contains("_КР_")
                    || doc.PathName.Contains("KR_")
                    || doc.PathName.Contains("_KR_"))
                {
                    if (elem is ElementType currentType)
                    {
                        WPFEntity typeNameError = CheckSysytemFamilyTypeName(currentType);
                        if (typeNameError != null)
                            result.Add(typeNameError);
                    }
                }
            }

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        private void FailuresProcessor(object sender, Autodesk.Revit.DB.Events.FailuresProcessingEventArgs e)
        {
            FailuresAccessor fAcc = e.GetFailuresAccessor();
            List<FailureMessageAccessor> failureMessageAccessors = fAcc.GetFailureMessages().ToList();
            if (failureMessageAccessors.Count > 0)
            {
                List<ElementId> elemsToDelete = new List<ElementId>();
                foreach (FailureMessageAccessor fma in failureMessageAccessors)
                {
                    Document fDoc = fAcc.GetDocument();

                    List<ElementId> fmFailElemsId = fma.GetFailingElementIds().ToList();
                    foreach (ElementId elId in fmFailElemsId)
                    {
                        Element fmFailElem = fDoc.GetElement(elId);
                        Type fmType = fmFailElem.GetType();
                        if (!fmType.Equals(typeof(PlanarFace))
                            && !fmType.Equals(typeof(ReferencePlane)))
                        {
                            elemsToDelete.Add(elId);
                        }
                    }
                }

                fAcc.DeleteAllWarnings();
                if (elemsToDelete.Count > 0)
                {
                    try
                    {
                        fAcc.DeleteElements(elemsToDelete);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException)
                    {
                        e.SetProcessingResult(FailureProcessingResult.Continue);
                        return;
                    }

                    e.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
                    return;
                }

                e.SetProcessingResult(FailureProcessingResult.Continue);
            }
        }

        /// <summary>
        /// Проверка имен семейства и его типоразмеров на наличие дубликатов
        /// </summary>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="docFamilies">Коллекция семейств проекта</param>
        private IEnumerable<WPFEntity> CheckFamilyAndTypeDuplicateName(Family currentFam, Element[] docFamilies)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            string currentFamName = currentFam.Name;
            if (Regex.Match(currentFamName, @"\b[.0]\d*$").Value.Length > 2)
            {
                result.Add(new WPFEntity(
                    ESEntity,
                    currentFam,
                    "Ошибка семейства",
                    "Данное семейство - это резервная копия. Запрещено использовать резервные копии!",
                    "Необходимо корректно обновить семейство. Резервные копии - могут содержать не корректную информацию.",
                    false));
            }

            string similarFamilyName = SearchSimilarName(currentFamName, docFamilies);
            if (!similarFamilyName.Equals(String.Empty))
            {
                result.Add(new WPFEntity(
                    ESEntity,
                    (Element)currentFam,
                    "Предупреждение семейства",
                    $"Возможно семейство является копией семейства \"{similarFamilyName}\"",
                    "Копий семейств в проекте быть не должно.",
                    false,
                    KPLN_ModelChecker_Lib.ErrorStatus.Warning));
            }

            ISet<ElementId> famSymolsIds = currentFam.GetFamilySymbolIds();
            FamilySymbol[] currentFamilySymols = new FamilySymbol[famSymolsIds.Count];
            for (int i = 0; i < famSymolsIds.Count; i++)
            {
                FamilySymbol symbol = currentFam.Document.GetElement(famSymolsIds.ElementAt(i)) as FamilySymbol;
                currentFamilySymols[i] = symbol;
            }

            foreach (FamilySymbol currentSymbol in currentFamilySymols)
            {
                string currentSymName = currentSymbol.Name;
                string similarSymbolName = SearchSimilarName(currentSymName, currentFamilySymols);

                if (!similarSymbolName.Equals(String.Empty))
                {
                    result.Add(new WPFEntity(
                        ESEntity,
                        (Element)currentFam,
                        "Предупреждение типоразмера",
                        $"Возможно тип является копией типоразмера \"{similarSymbolName}\"",
                        "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!",
                        false,
                        KPLN_ModelChecker_Lib.ErrorStatus.Warning));
                }
            }

            return result;
        }

        /// <summary>
        /// Проверка имен типоразмеров системных семейств
        /// </summary>
        /// <param name="elemType">Тип для проверки</param>
        private WPFEntity CheckSysytemFamilyTypeName(ElementType elemType)
        {
            string typeName = elemType.Name;
            if (typeName.Equals("99_Не использовать"))
                return null;

            string[] typeSplitedName = typeName.Split('_');
            if (typeSplitedName.Length < 3)
            {
                return new WPFEntity(
                    ESEntity,
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Данный типоразмер назван не по ВЕР - не хватает основных блоков",
                    "Имя системных типоразмеров делиться минимум на 3 блока: код, шифр слоёв и описание. Разделитель - нижнее подчеркивание '_'",
                    false,
                    true);
            }

            if (!(typeSplitedName[0].StartsWith("00")
                || typeSplitedName[0].StartsWith("01")
                || typeSplitedName[0].StartsWith("02")
                || typeSplitedName[0].StartsWith("03")
                || typeSplitedName[0].StartsWith("04")
                || typeSplitedName[0].StartsWith("05")))
            {
                return new WPFEntity(
                    ESEntity,
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Данный типоразмер назван не по ВЕР - ошибка кода",
                    "Имя системных типоразмеров может иметь коды: 00, 01, 02, 03, 04, 05.",
                    false,
                    true);
            }

            #region Проверка ЖБ на привязку к коду 00
            string sliceCode = typeSplitedName[1];
            if (sliceCode.ToUpper().Equals("ВН") || sliceCode.ToUpper().Equals("НА"))
                sliceCode = typeSplitedName[2];

            if (typeSplitedName[0].Equals("00") && !sliceCode.ToUpper().Contains("ЖБ") && !sliceCode.ToUpper().StartsWith("К"))
            {
                return new WPFEntity(
                    ESEntity,
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Код '00_' может содержать только несущие конструкции",
                    $"Несущий стены/перекрытия - это ЖБ, К (для стен) (аббревиатуры указаны в ВЕР). Сейчас аббревиатура не содержит бетон, или кирпич (нет ЖБ/К): \"{sliceCode}\"",
                    false,
                    true);
            }
            if (sliceCode.ToUpper().Contains("ЖБ") && !typeSplitedName[0].Equals("00"))
            {
                return new WPFEntity(
                    ESEntity,
                    elemType,
                    "Предупреждение типоразмера системного",
                    $"ЖБ вне несущего слоя",
                    $"Скорее всего это ошибка, т.к. ЖБ используется вне несущего слоя (код не 00, а \"{typeSplitedName[0]}\")",
                    false,
                    ErrorStatus.Warning,
                    true);
            }
            #endregion

            #region Нахожу суммарную толщину
            string totalThicknessStr = typeSplitedName[typeSplitedName.Length - 1];
            if (!double.TryParse(totalThicknessStr, out double totalThickness))
            {
                totalThicknessStr = typeSplitedName[typeSplitedName.Length - 2];
                if (!double.TryParse(typeSplitedName[typeSplitedName.Length - 2], out totalThickness))
                {
                    totalThicknessStr = typeSplitedName[typeSplitedName.Length - 3];
                    if (!double.TryParse(typeSplitedName[typeSplitedName.Length - 3], out totalThickness))
                    {
                        return new WPFEntity(
                            ESEntity,
                            elemType,
                            "Ошибка типоразмера системного",
                            "Ошибка индекса положения суммарной толщины",
                            $"Толщина слоя указывается в последнем, или предпоследнем блоке имени типоразмера. Блоки имен разделяются нижним подчеркиванием \"_\". " +
                                $"Сейчас это место занимамет не цифра, а: \"{totalThicknessStr}\". Нужно исправить имя типа в соотвествии с ВЕР.",
                            false,
                            true);
                    }
                }
            }

            double typeThickness = 0;
            if (elemType is FloorType floorType)
#if Revit2020 || Debug2020
                typeThickness = UnitUtils.ConvertFromInternalUnits(floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble(),
                    DisplayUnitType.DUT_MILLIMETERS);
#endif
#if Revit2023 || Debug2023
                typeThickness = UnitUtils.ConvertFromInternalUnits(floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble(),
                        new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
            else if (elemType is WallType wallType)
            {
                Parameter widthParam = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                if (widthParam == null)
                    return null;
#if Revit2020 || Debug2020
                typeThickness = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(),
                    DisplayUnitType.DUT_MILLIMETERS);
#endif
#if Revit2023 || Debug2023
                    typeThickness = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(),
                        new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
            }

            if (Math.Abs(totalThickness - typeThickness) > 0.1)
            {
                return new WPFEntity(
                    ESEntity,
                    elemType,
                    "Ошибка типоразмера системного",
                    "Сумма слоёв не совпадает с описанием",
                    $"Толщина слоя в имени указана как \"{totalThicknessStr}\", хотя на самом деле она составляет \"{typeThickness}\"",
                    false,
                    true);
            }
            #endregion

            return null;
        }

        /// <summary>
        /// Проверка пути к семейству
        /// </summary>
        /// <param name="doc">Файл Revit</param>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="outputCollection">Коллекция элементов WPFDisplayItem для отчета</param>
        private WPFEntity CheckFamilyPath(Document doc, Family currentFam)
        {
            // Отсеиваю по имени семейств плагинов
            if (currentFam.Name.Contains("ClashPoint")
                || currentFam.Name.Contains("Определения для Weandrevit ALL"))
                return null;

            // Блок игнорирования семейств ostec/dkc (они плагином устанавливаются локально на диск С)
            if (currentFam.Name.ToLower().Contains("ostec")
                || currentFam.Name.ToLower().Contains("dkc"))
                return null;

            Category currentCat = currentFam.FamilyCategory;
            if (currentCat == null)
                return null;

            // Блок игнорирования семейств настроенных из шаблона (АР балясины, ограждения)
            if (currentCat.Id.IntegerValue == (int)BuiltInCategory.OST_StairsRailingBaluster
                || currentCat.Id.IntegerValue == (int)BuiltInCategory.OST_RailingTermination
                || currentCat.Id.IntegerValue == (int)BuiltInCategory.OST_RailingSupport)
                return null;

            // Блок игнорирования семейств аннотаций, кроме штампов (остальное проектировщики могут создавать)
            if (currentCat.CategoryType.Equals(CategoryType.Annotation)
                && !currentFam.Name.StartsWith("020_")
                && !currentFam.Name.StartsWith("022_")
                && !currentFam.Name.StartsWith("023_")
                && !currentFam.Name.ToLower().Contains("жук"))
                return null;

            BuiltInCategory currentBIC = (BuiltInCategory)currentCat.Id.IntegerValue;
            if (currentFam.get_Parameter(BuiltInParameter.FAMILY_SHARED).AsInteger() != 1
                && currentFam.IsEditable
                && !currentBIC.Equals(BuiltInCategory.OST_ProfileFamilies)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponents)
                && !currentBIC.Equals(BuiltInCategory.OST_GenericAnnotation)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponentsHiddenLines)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponentTags))
            {
                Document famDoc;
                try
                {
                    famDoc = doc.EditFamily(currentFam);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Работа остановлена, т.к. семейство {currentFam.Name} не может быть открыто. Причина: {ex}");
                }
                if (famDoc.IsFamilyDocument != true)
                    return null;

                string famPath = famDoc.PathName;
                if (!(famPath.StartsWith("X:\\") && new FileInfo(famPath).Exists)
                    & !famPath.Contains("03_Скрипты")
                    & !famPath.Contains("KPLN_Loader"))
                {
                    return new WPFEntity(
                        ESEntity,
                        currentFam,
                        "Ошибка источника семейства",
                        "Данное семейство - не с диска Х. Запрещено использовать сторонние источники!",
                        "Использовать в проекте данное семейство можно только по согласованию в BIM-отделе.",
                        false);
                }

                famDoc.Close(false);
            }

            return null;
        }

        /// <summary>
        /// Поиск похожего имени. Одинаковым должна быть только первичная часть имени, до среза по циферным значениям
        /// </summary>
        /// <param name="currentName">Имя, которое нужно проанализировать</param>
        /// <param name="elemsColl">Коллекция, по которой нужно осуществлять поиск</param>
        /// <returns>Имя подобного элемента</returns>
        private string SearchSimilarName(string currentName, Element[] elemsColl)
        {
            string similarFamilyName = String.Empty;

            // Осуществляю поиск цифр в конце имени
            string digitEndTrimmer = Regex.Match(currentName, @"\d*$").Value;
            // Осуществляю срез имени на найденные цифры в конце имени
            string truePartOfName = currentName.TrimEnd(digitEndTrimmer.ToArray());
            if (digitEndTrimmer.Length > 0)
            {
                foreach (Element checkElem in elemsColl)
                {
                    if (!checkElem.Equals(currentName) && checkElem.Name.Equals(truePartOfName.TrimEnd(new char[] { ' ' })))
                    {
                        similarFamilyName = checkElem.Name;
                    }
                }
            }
            return similarFamilyName;
        }
    }
}

using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckFamilies : AbstrCheck
    {
        public CheckFamilies() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка семейств";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CommandCheckFamilies",
                    new Guid("168c83b9-1d62-4d3f-9bbb-fd1c1e9a0807"),
                    new Guid("168c83b9-1d62-4d3f-9bbb-fd1c1e9a0808"));
        }

        public override Element[] GetElemsToCheck()
        {
            FilteredElementCollector docFamsColl = new FilteredElementCollector(CheckDocument).OfClass(typeof(Family));
            FilteredElementCollector wallTypesColl = new FilteredElementCollector(CheckDocument).OfClass(typeof(WallType));
            FilteredElementCollector floorTypesColl = new FilteredElementCollector(CheckDocument).OfClass(typeof(FloorType));

            return docFamsColl
                .UnionWith(wallTypesColl)
                .UnionWith(floorTypesColl)
                .ToArray();
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            CheckUIApp.Application.FailuresProcessing += FailuresProcessor;

            try
            {
                foreach (Element elem in elemColl)
                {
                    // Проверяю семейства и их типоразмеры
                    if (elem is Family currentFam)
                    {
                        _checkerEntitiesCollHeap.AddRange(CheckFamilyAndTypeDuplicateName(currentFam, elemColl));

                        // Проверка пути семейства - ТОЛЬКО для спецов BIM-отдела
                        if (DBMainService.CurrentDBUser.SubDepartmentId == 8)
                        {
                            CheckerEntity checkFamilyPath = CheckFamilyPath(CheckDocument, currentFam);
                            if (checkFamilyPath != null)
                                _checkerEntitiesCollHeap.Add(checkFamilyPath);
                        }
                    }

                    // Проверяю системные типоразмеры АР и КР
                    else if (CheckDocument.PathName.Contains("АР_")
                        || CheckDocument.PathName.Contains("_АР_")
                        || CheckDocument.PathName.Contains("AR_")
                        || CheckDocument.PathName.Contains("_AR_")
                        || CheckDocument.PathName.Contains("КР_")
                        || CheckDocument.PathName.Contains("_КР_")
                        || CheckDocument.PathName.Contains("KR_")
                        || CheckDocument.PathName.Contains("_KR_"))
                    {
                        if (elem is ElementType currentType)
                        {
                            CheckerEntity typeNameError = CheckSysytemFamilyTypeName(currentType);
                            if (typeNameError != null)
                                _checkerEntitiesCollHeap.Add(typeNameError);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Дополнительная обертка из дополнительного try/cath ради отписки от события FailuresProcessor. Ошибку словит try/cath из метода ExecuteCheck
                throw ex;
            }
            finally
            {
                CheckUIApp.Application.FailuresProcessing -= FailuresProcessor;
            }

            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Проверка пути к семейству
        /// </summary>
        /// <param name="doc">Файл Revit</param>
        /// <param name="currentFam">Семейство для проверки</param>
        private CheckerEntity CheckFamilyPath(Document doc, Family currentFam)
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
                    return new CheckerEntity(
                        currentFam,
                        "Ошибка источника семейства (только для BIM-отдела)",
                        "Данное семейство - не с диска Х. Запрещено использовать сторонние источники!",
                        "ВАЖНО: Данную проверку запускают только специалисты BIM-отдела. " +
                            "Ошибку выдавать исполнителям ТОЛЬКО если семейство не совпадает с шаблоном моделирования, и требует серьёзных правок");
                }

                famDoc.Close(false);
            }

            return null;
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
        private IEnumerable<CheckerEntity> CheckFamilyAndTypeDuplicateName(Family currentFam, Element[] docFamilies)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            string currentFamName = currentFam.Name;
            if (Regex.Match(currentFamName, @"\b[.0]\d*$").Value.Length > 2)
            {
                result.Add(new CheckerEntity(
                    currentFam,
                    "Ошибка семейства",
                    "Данное семейство - это резервная копия. Запрещено использовать резервные копии!",
                    "Необходимо корректно обновить семейство. Резервные копии - могут содержать не корректную информацию."));
            }

            string similarFamilyName = SearchSimilarName(currentFamName, docFamilies);
            if (!similarFamilyName.Equals(String.Empty))
            {
                result.Add(new CheckerEntity(
                    (Element)currentFam,
                    "Предупреждение семейства",
                    $"Возможно семейство является копией семейства \"{similarFamilyName}\"",
                    "Копий семейств в проекте быть не должно.")
                    .Set_Status(ErrorStatus.Warning));
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
                    result.Add(new CheckerEntity(
                        (Element)currentSymbol,
                        "Предупреждение типоразмера",
                        $"Возможно тип является копией типоразмера \"{similarSymbolName}\"",
                        "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!")
                        .Set_Status(ErrorStatus.Warning));
                }
            }

            return result;
        }

        /// <summary>
        /// Проверка имен типоразмеров системных семейств
        /// </summary>
        /// <param name="elemType">Тип для проверки</param>
        private CheckerEntity CheckSysytemFamilyTypeName(ElementType elemType)
        {
            string typeName = elemType.Name;
            if (typeName.Equals("99_Не использовать"))
                return null;

            string[] typeSplitedName = typeName.Split('_');
            if (typeSplitedName.Length < 3)
            {
                return new CheckerEntity(
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Данный типоразмер назван не по ВЕР - не хватает основных блоков",
                    "Имя системных типоразмеров делиться минимум на 3 блока: код, шифр слоёв и описание. Разделитель - нижнее подчеркивание '_'")
                    .Set_CanApprovedAndESData(ESEntity);
            }

            if (!(typeSplitedName[0].StartsWith("00")
                || typeSplitedName[0].StartsWith("01")
                || typeSplitedName[0].StartsWith("02")
                || typeSplitedName[0].StartsWith("03")
                || typeSplitedName[0].StartsWith("04")
                || typeSplitedName[0].StartsWith("05")))
            {
                return new CheckerEntity(
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Данный типоразмер назван не по ВЕР - ошибка кода",
                    "Имя системных типоразмеров может иметь коды: 00, 01, 02, 03, 04, 05.")
                    .Set_CanApprovedAndESData(ESEntity);
            }

            #region Проверка ЖБ на привязку к коду 00
            string sliceCode = typeSplitedName[1];
            if (sliceCode.ToUpper().Equals("ВН") || sliceCode.ToUpper().Equals("НА"))
                sliceCode = typeSplitedName[2];

            if (typeSplitedName[0].Equals("00") && !sliceCode.ToUpper().Contains("ЖБ") && !sliceCode.ToUpper().StartsWith("К"))
            {
                return new CheckerEntity(
                    elemType,
                    "Ошибка типоразмера системного",
                    $"Код '00_' может содержать только несущие конструкции",
                    $"Несущий стены/перекрытия - это ЖБ, К (для стен) (аббревиатуры указаны в ВЕР). Сейчас аббревиатура не содержит бетон, или кирпич (нет ЖБ/К): \"{sliceCode}\"")
                    .Set_CanApprovedAndESData(ESEntity);
            }
            if (sliceCode.ToUpper().Contains("ЖБ") && !typeSplitedName[0].Equals("00"))
            {
                return new CheckerEntity(
                    elemType,
                    "Предупреждение типоразмера системного",
                    $"ЖБ вне несущего слоя",
                    $"Скорее всего это ошибка, т.к. ЖБ используется вне несущего слоя (код не 00, а \"{typeSplitedName[0]}\")")
                    .Set_CanApprovedAndESData(ESEntity, ErrorStatus.Warning);
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
                        return new CheckerEntity(
                            elemType,
                            "Ошибка типоразмера системного",
                            "Ошибка индекса положения суммарной толщины",
                            $"Толщина слоя указывается в последнем, или предпоследнем блоке имени типоразмера. Блоки имен разделяются нижним подчеркиванием \"_\". " +
                                $"Сейчас это место занимамет не цифра, а: \"{totalThicknessStr}\". Нужно исправить имя типа в соотвествии с ВЕР.")
                            .Set_CanApprovedAndESData(ESEntity);
                    }
                }
            }

            double typeThickness = 0;
            if (elemType is FloorType floorType)
#if Revit2020 || Debug2020
                typeThickness = UnitUtils.ConvertFromInternalUnits(floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM).AsDouble(),
                    DisplayUnitType.DUT_MILLIMETERS);
#endif
#if !Revit2020 && !Debug2020
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
#if !Revit2020 && !Debug2020
                typeThickness = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(),
                    new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
            }

            if (Math.Abs(totalThickness - typeThickness) > 0.1)
            {
                return new CheckerEntity(
                    elemType,
                    "Ошибка типоразмера системного",
                    "Сумма слоёв не совпадает с описанием",
                    $"Толщина слоя в имени указана как \"{totalThicknessStr}\", хотя на самом деле она составляет \"{typeThickness}\"")
                    .Set_CanApprovedAndESData(ESEntity);
            }
            #endregion

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

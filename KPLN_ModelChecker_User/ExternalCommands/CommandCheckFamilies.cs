using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckFamilies : AbstrCheckCommand<CommandCheckFamilies>, IExternalCommand
    {
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
            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Application app = uiapp.Application;
            app.FailuresProcessing += FailuresProcessor;

            try
            {
                // Получаю коллекцию элементов для анализа
                Element[] famColl = new FilteredElementCollector(doc).OfClass(typeof(Family)).ToArray();

                #region Проверяю и обрабатываю элементы
                WPFEntity[] wpfColl = CheckCommandRunner(doc, famColl);
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

            foreach (Family currentFam in elemColl)
            {
                result.AddRange(CheckFamilyName(currentFam, elemColl));
                WPFEntity checkFamilyPath = CheckFamilyPath(doc, currentFam);
                if (checkFamilyPath != null)
                    result.Add(checkFamilyPath);
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
                    //elemsToDelete.AddRange(fma.GetFailingElementIds());

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
        /// Проверка имен семейства и его типоразмеров
        /// </summary>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="docFamilies">Коллекция семейств проекта</param>
        /// <param name="outputCollection">Коллекция элементов WPFDisplayItem для отчета</param>
        private IEnumerable<WPFEntity> CheckFamilyName(Family currentFam, Element[] docFamilies)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            string currentFamName = currentFam.Name;
            if (Regex.Match(currentFamName, @"\b[.0]\d*$").Value.Length > 2)
            {
                result.Add(new WPFEntity(
                    currentFam,
                    Status.Error,
                    "Предупреждение семейства",
                    $"Данное семейство - это резервная копия. Запрещено использовать резервные копии!",
                    false,
                    false,
                    "Необходимо корректно обновить семейство. Резервные копии - могут содержать не корректную информацию."));
            }

            string similarFamilyName = SearchSimilarName(currentFamName, docFamilies);
            if (!similarFamilyName.Equals(String.Empty))
            {
                result.Add(new WPFEntity(
                    currentFam,
                    Status.Warning,
                    "Предупреждение семейства",
                    $"Возможно семейство является копией семейства «{similarFamilyName}»",
                    false,
                    false,
                    "Копий семейств в проекте быть не должно."));
            }

            ISet<ElementId> famSymolsIds = currentFam.GetFamilySymbolIds();
            Element[] currentFamilySymols = new Element[famSymolsIds.Count];
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
                    currentFam,
                    Status.Warning,
                    "Предупреждение типоразмера",
                    $"Возможно тип является копией типоразмера «{similarSymbolName}»",
                    false,
                    false,
                    "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!"));
                }
            }

            return result;
        }

        /// <summary>
        /// Проверка пути к семейству
        /// </summary>
        /// <param name="doc">Файл Revit</param>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="outputCollection">Коллекция элементов WPFDisplayItem для отчета</param>
        private WPFEntity CheckFamilyPath(Document doc, Family currentFam)
        {
            BuiltInCategory currentBIC;
            Category currentCat = currentFam.FamilyCategory;
            if (currentCat == null)
                return null;

            currentBIC = (BuiltInCategory)currentCat.Id.IntegerValue;
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

                // Блок игнорирования семейств ostec/dkc (они плагином устанавливаются локально на диск С)
                if (currentFam.Name.ToLower().Contains("ostec")
                    || currentFam.Name.ToLower().Contains("dkc"))
                    return null;

                // Блок игнорирования семейств аннотаций, кроме штампов (остальное проектировщики могут создавать)
                if (currentCat.CategoryType.Equals(CategoryType.Annotation)
                    && !currentFam.Name.StartsWith("020_")
                    && !currentFam.Name.StartsWith("022_")
                    && !currentFam.Name.StartsWith("023_")
                    && !currentFam.Name.ToLower().Contains("жук"))
                    return null;

                string famPath = famDoc.PathName;
                if (!famPath.StartsWith("X:\\")
                    & !famPath.Contains("03_Скрипты")
                    & !famPath.Contains("KPLN_Loader"))
                {
                    return new WPFEntity(
                        currentFam,
                        Status.Error,
                        "Предупреждение источника семейства",
                        $"Данное семейство - не с диска Х. Запрещено использовать сторонние источники!",
                        false,
                        false,
                        "Использовать в проекте данное семейство можно только по согласованию в BIM-отделе.");
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

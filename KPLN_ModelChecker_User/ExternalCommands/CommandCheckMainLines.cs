using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    internal struct GridEntity
    {
        internal Grid CurrentGrid {  get; set; }
        
        internal Grid NearestGrid {  get; set; }

        internal double Distance { get; set; }

        internal bool IsChecked { get; set; }
        
        internal bool IsDistanceError{ get; set; }
    }
    
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckMainLines : AbstrCheckCommand<CommandCheckMainLines>, IExternalCommand
    {
        internal const string PluginName = "Проверка осей/уровней";

        public CommandCheckMainLines() : base()
        {
        }

        internal CommandCheckMainLines(ExtensibleStorageEntity esEntity) : base(esEntity)
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

            // Получаю коллекцию элементов для анализа
            IEnumerable<Element> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType();
            IEnumerable<Element> grids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();
            Element[] mainLInesColl = levels.Union(grids).ToArray();

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, mainLInesColl);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();
            foreach (Element element in elemColl)
            {
                WPFEntity monitorErrorEnt = GetMonitoringErrorEnt(ESEntity, _uiApp, element);
                if (monitorErrorEnt !=  null)
                    result.Add(monitorErrorEnt);

                WPFEntity pinErrorEnt = GetPinErrorEnt(ESEntity, element);
                if (pinErrorEnt != null)
                    result.Add(pinErrorEnt);
            }

            Grid[] lineGridsOnly = elemColl
                .Where(el => el is Grid gr && gr.Curve is Line)
                .Cast<Grid>()
                .OrderBy(gr => gr.Name)
                .ToArray();

            IEnumerable<WPFEntity> gridParallDistErrEnts = GetGridParallDistErrorEnt(ESEntity, lineGridsOnly);
            if (gridParallDistErrEnts != null)
                result.AddRange(gridParallDistErrEnts);

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Проверка на наличие мониторинга
        /// </summary>
        /// <returns></returns>
        private static WPFEntity GetMonitoringErrorEnt(ExtensibleStorageEntity esEntity, UIApplication uiapp, Element element)
        {
            if (element.IsMonitoringLinkElement())
            {
                foreach (ElementId i in element.GetMonitoredLinkElementIds())
                {
                    if (!(uiapp.ActiveUIDocument.Document.GetElement(i) is RevitLinkInstance link))
                    {
                        return new WPFEntity(
                            esEntity,
                            element,
                            "Ошибка мониторинга",
                            $"Связь не найдена: «{element.Name}»",
                            $"Элементу с ID {element.Id} необходимо исправить мониторинг",
                            false);
                    }
                    else if (!link.Name.ToLower().Contains("разб")
                        && !link.Name.Contains("СЕТ_1_1-3_00_РФ"))
                    {
                        return new WPFEntity(
                            esEntity,
                            element,
                            "Ошибка мониторинга",
                            $"Мониторинг не из разбивочного файла: «{element.Name}»",
                            $"Элементу с ID {element.Id} необходимо исправить мониторинг, сейчас он присвоен связи {link.Name}",
                            false);
                    }
                }

                return null;
            }
            else
            {
                return new WPFEntity(
                    esEntity,
                    element,
                    "Отсутсвие мониторинга",
                    $"Элементу с ID {element.Id} необходимо задать мониторинг",
                    string.Empty,
                    false);
            }
        }

        /// <summary>
        /// Проверка на прикрепление
        /// </summary>
        /// <param name="esEntity"></param>
        /// <param name="uiapp"></param>
        /// <param name="element"></param>
        /// <returns></returns>
        private static WPFEntity GetPinErrorEnt(ExtensibleStorageEntity esEntity, Element element)
        {
            if (element.Pinned) return null;

            return new WPFEntity(
                esEntity,
                element,
                "Отсутсвие прикрепления (PIN)",
                $"Элемент не прикреплен: «{element.Name}»",
                $"Элемент с ID {element.Id} необходимо прикрепить (PIN)",
                false);
        }

        /// <summary>
        /// Проверка осей на расстояние между параллельными эл-тами (ТОЛЬКО прямые оси)
        /// </summary>
        /// <param name="esEntity"></param>
        /// <param name="element"></param>
        /// <returns></returns>
        private static IEnumerable<WPFEntity> GetGridParallDistErrorEnt(ExtensibleStorageEntity esEntity, Grid[] grids)
        {
            GridEntity[] gridEntities = grids
                .Select(gr => new GridEntity { CurrentGrid = gr} )
                .ToArray();

            for (int i = 0; i < gridEntities.Count(); i++)
            {
                GridEntity gridEntity = gridEntities[i];
                GridEntity nearestGridEntity = gridEntities.FirstOrDefault(ge => ge.NearestGrid != null && ge.NearestGrid.Id == gridEntity.CurrentGrid.Id);
                if (nearestGridEntity.IsChecked && nearestGridEntity.IsDistanceError) continue;

                SetGridData(grids, ref gridEntities[i]);
            }

            return gridEntities
                .Where(ge => ge.IsDistanceError)
                .Select(ge => new WPFEntity
                    (esEntity,
                    ge.CurrentGrid,
                    "Нарушена точность построения",
                    $"Ось «{ge.CurrentGrid.Name}» расположена на расстоянии {ge.Distance} от оси «{ge.NearestGrid.Name}»",
                    $"Точность построения осей - 5,00001. ПРЕДВАРИТЕЛЬНО - согласуй внесение изменений со своим BIM-координатором",
                    false,
                    true));
        }

        /// <summary>
        /// Устанавливаю поля GridEntity анализируя ПАРАЛЛЕЛЬНЫЕ оси
        /// </summary>
        /// <param name="gridEnt"></param>
        /// <param name="grids"></param>
        private static void SetGridData(Grid[] grids, ref GridEntity gridEnt)
        {
            Grid checkGrid = gridEnt.CurrentGrid;

            Curve checkGridCurve = checkGrid.Curve;
            XYZ fEndPontCheckGrid = checkGridCurve.GetEndPoint(0);
            XYZ sEndPontCheckGrid = checkGridCurve.GetEndPoint(1);

            Line checkGridLine = checkGridCurve as Line;
            XYZ checkGridDirection = checkGridLine.Direction;

            Grid tempNearestGrid = null;
            double tempDist = double.MaxValue;
            foreach (Grid grid in grids)
            {
                if (checkGrid.Id == grid.Id) continue;

                Curve gridCurve = grid.Curve;
                XYZ fStartPontGrid = gridCurve.GetEndPoint(0);
                XYZ sStartPontGrid = gridCurve.GetEndPoint(1);

                Line lineGrid = grid.Curve as Line;
                XYZ gridDirection = lineGrid.Direction;

                // Проверяем на параллельность (если векторное произведение != 0, то линии НЕ параллельны, пропускаем)
                XYZ crossProduct = checkGridDirection.CrossProduct(gridDirection);
                if (!crossProduct.IsZeroLength()) continue;


                // Расстояние между параллельными линиями
                IntersectionResult intRes1 = checkGrid.Curve.Project(fStartPontGrid);
                IntersectionResult intRes2 = grid.Curve.Project(fEndPontCheckGrid);
                double distBetweenIntResPnts1 = intRes1.XYZPoint.DistanceTo(intRes2.XYZPoint);

                IntersectionResult intRes3 = checkGrid.Curve.Project(fStartPontGrid);
                IntersectionResult intRes4 = grid.Curve.Project(sEndPontCheckGrid);
                double distBetweenIntResPnts3 = intRes3.XYZPoint.DistanceTo(intRes4.XYZPoint);

                double minBetweenDist = Math.Min(distBetweenIntResPnts1, distBetweenIntResPnts3);

#if Debug2020 || Revit2020
                double resultDistMM = UnitUtils.ConvertFromInternalUnits(minBetweenDist, DisplayUnitType.DUT_MILLIMETERS);
#else
                double resultDistMM = UnitUtils.ConvertFromInternalUnits(minBetweenDist, UnitTypeId.Millimeters);
#endif

                if (resultDistMM < tempDist)
                {
                    tempDist = resultDistMM;
                    tempNearestGrid = grid;
                }
            }

            // Предварительно округляю с нужно точностью, чтобы и в отчет попало норм. значение, и анализировалось тоже корректное значение, а не напрм. 14.999999975
            double roundDist= Math.Round(tempDist, 5);

            gridEnt.NearestGrid = tempNearestGrid;
            gridEnt.Distance = roundDist;
            gridEnt.IsChecked = true;
            gridEnt.IsDistanceError = IsErrorValidNumber(roundDist);
        }

        private static bool IsErrorValidNumber(double roundNumber)
        {
            string roundNumberString = roundNumber.ToString("F10").TrimEnd('0');
            char separ = ',';
            if (!roundNumberString.Contains(separ))
                separ = '.';

            string roundNumberIntPart = roundNumberString.Split(separ)[0];
            int decimalPlaces = roundNumberString.Contains(separ) ? roundNumberString.Split(separ)[1].Length : 0;

            if (int.TryParse(roundNumberIntPart, out int intPart))
            {
                bool isCorrectEnding = intPart % 10 == 0 || intPart % 10 == 5;
                bool isValidPrecision = decimalPlaces == 0 || decimalPlaces > 5;
                
                return !(isCorrectEnding && isValidPrecision);
            }

            return false;
        }
    }
}

using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    internal struct GridEntity
    {
        internal Grid CurrentGrid { get; set; }

        internal Grid NearestParallelGrid { get; set; }

        internal double NearestParallelDistance { get; set; }

        internal Grid NearestAnglelGrid { get; set; }

        internal double NearestAngle { get; set; }

        internal bool IsChecked { get; set; }

        internal bool IsDistanceError { get; set; }

        internal bool IsAngleError { get; set; }
    }

    public sealed class CheckMainLines : AbstrCheck
    {
        public CheckMainLines() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка осей/уровней";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CommandCheckMainLines",
                    new Guid("eac2c205-342d-4ba3-98a1-d82c82a4638e"));
        }


        public override Element[] GetElemsToCheck()
        {
            IEnumerable<Element> levels = new FilteredElementCollector(CheckDocument).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType();
            IEnumerable<Element> grids = new FilteredElementCollector(CheckDocument).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();

            return levels.Union(grids).ToArray();
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            string docTitle = CheckDocument.Title;
            bool isRazbFile = docTitle.Contains("_РФ_") || docTitle.Contains("Разб");
            foreach (Element element in elemColl)
            {
                if (!isRazbFile)
                {
                    CheckerEntity monitorErrorEnt = GetMonitoringErrorEnt(CheckDocument, element);
                    if (monitorErrorEnt != null)
                        _checkerEntitiesCollHeap.Add(monitorErrorEnt);
                }

                CheckerEntity pinErrorEnt = GetPinErrorEnt(element);
                if (pinErrorEnt != null)
                    _checkerEntitiesCollHeap.Add(pinErrorEnt);
            }

            Grid[] lineGridsOnly = elemColl
                .Where(el => el is Grid gr && gr.Curve is Line)
                .Cast<Grid>()
                .OrderBy(gr => gr.Name)
                .ToArray();

            if (lineGridsOnly.Length > 0)
            {
                GridEntity[] gridEntites = CreateGridEntites(lineGridsOnly);

                IEnumerable <CheckerEntity> gridParallDistErrEnts = GetGridParallDistErrorEnt(gridEntites);
                if (gridParallDistErrEnts != null)
                    _checkerEntitiesCollHeap.AddRange(gridParallDistErrEnts);

                // Проверка угла не реализуема. Причина - угол может быть не точным, при условии, что точки пересечения осей точные.
                //IEnumerable<CheckerEntity> gridAngleErrEnts = GetGridAngleErrorEnt(gridEntites);
                //if (gridAngleErrEnts != null)
                //    _checkerEntitiesCollHeap.AddRange(gridAngleErrEnts);
            }

            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Проверка на наличие мониторинга
        /// </summary>
        /// <returns></returns>
        private static CheckerEntity GetMonitoringErrorEnt(Document doc, Element element)
        {
            if (element.IsMonitoringLinkElement())
            {
                foreach (ElementId i in element.GetMonitoredLinkElementIds())
                {
                    if (!(doc.GetElement(i) is RevitLinkInstance link))
                    {
                        return new CheckerEntity(
                            element,
                            "Ошибка мониторинга",
                            $"Связь не найдена",
                            $"Мониторинг назначен на связь, которой больше не существует в модели.");
                    }
                    else if (!link.Name.ToLower().Contains("разб")
                        && !link.Name.Contains("СЕТ_1_1-3_00_РФ"))
                    {
                        return new CheckerEntity(
                            element,
                            "Ошибка мониторинга",
                            $"Мониторинг не из разбивочного файла",
                            $"Мониторинг может быть только из разбивочного файла, сейчас он присвоен связи {link.Name}");
                    }
                }

                return null;
            }
            else
            {
                return new CheckerEntity(
                    element,
                    "Отсутствие мониторинга",
                    $"Оси и уровни обязательно должны иметь мониторинг из разбивочного файла",
                    string.Empty);
            }
        }

        /// <summary>
        /// Проверка на прикрепление
        /// </summary>
        /// <returns></returns>
        private static CheckerEntity GetPinErrorEnt(Element element)
        {
            if (element.Document.GetElement(element.GroupId) is Group group)
            {
                if (group.Pinned || element.Pinned)
                    return null;
            }
            else if (element.Pinned)
                return null;

            return new CheckerEntity(
                element,
                "Отсутствие прикрепления (PIN)",
                $"Элемент не прикреплен",
                $"Элемент необходимо прикрепить (PIN). Если он в группе - достаточно прикрепить (PIN) группу");
        }

        /// <summary>
        /// Создание коллекции GridEntity для анализа (ТОЛЬКО прямые оси)
        /// </summary>
        /// <param name="grids"></param>
        /// <returns></returns>
        private static GridEntity[] CreateGridEntites(Grid[] grids)
        {
            GridEntity[] gridEntities = grids
                .Select(gr => new GridEntity { CurrentGrid = gr })
                .ToArray();

            for (int i = 0; i < gridEntities.Count(); i++)
            {
                GridEntity gridEntity = gridEntities[i];
                GridEntity nearestGridEntity = gridEntities.FirstOrDefault(ge => ge.NearestParallelGrid != null && ge.NearestParallelGrid.Id == gridEntity.CurrentGrid.Id);
                if (nearestGridEntity.IsChecked && (nearestGridEntity.IsDistanceError || nearestGridEntity.IsAngleError)) continue;

                SetGridData(grids, ref gridEntities[i]);
            }

            return gridEntities;
        }



        /// <summary>
        /// Сортировка GridEntity по расстоянию между осями
        /// </summary>
        /// <returns></returns>
        private static IEnumerable<CheckerEntity> GetGridParallDistErrorEnt(GridEntity[] gridEntities)
        {
            return gridEntities
                .Where(ge => ge.IsDistanceError)
                .Select(ge => new CheckerEntity(new Element[2] { ge.CurrentGrid, ge.NearestAnglelGrid },
                    "Нарушена точность построения",
                    $"Ось «{ge.CurrentGrid.Name}» расположена на расстоянии {ge.NearestParallelDistance} от оси «{ge.NearestParallelGrid.Name}»",
                    $"Точность построения осей - 5,00001. ПРЕДВАРИТЕЛЬНО - согласуй внесение изменений со своим BIM-координатором"));
        }

        /// <summary>
        /// Сортировка GridEntity по углу между ними
        /// </summary>
        /// <returns></returns>
        private static IEnumerable<CheckerEntity> GetGridAngleErrorEnt(GridEntity[] gridEntities)
        {
            return gridEntities
                .Where(ge => ge.IsAngleError)
                .Select(ge => new CheckerEntity(new Element[2] { ge.CurrentGrid, ge.NearestAnglelGrid },
                    "Нарушена точность построения",
                    $"Ось «{ge.CurrentGrid.Name}» расположена под углом {ge.NearestAngle}° от оси «{ge.NearestAnglelGrid.Name}»",
                    "Точность построения угловых размеров осей - 0,00001°. ПРЕДВАРИТЕЛЬНО - согласуй внесение изменений со своим BIM-координатором"));
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
            Grid tempNearestAnglelGrid = null;
            double tempDist = double.MaxValue;
            double tempAngleDegrees = 0;
            foreach (Grid grid in grids)
            {
                // Исключение проверки сам на себя
                if (checkGrid.Id == grid.Id) continue;

                Curve gridCurve = grid.Curve;
                XYZ fStartPontGrid = gridCurve.GetEndPoint(0);
                XYZ sStartPontGrid = gridCurve.GetEndPoint(1);

                Line lineGrid = gridCurve as Line;
                XYZ gridDirection = lineGrid.Direction;

                // Проверяем на параллельность. Еесли векторное произведение != 0, то линии НЕ параллельны, проверяем угол
                XYZ crossProduct = checkGridDirection.CrossProduct(gridDirection);
                if (!crossProduct.IsZeroLength())
                {
                    // Проверка угла не реализуема. Причина - угол может быть не точным, при условии, что точки пересечения осей точные.
                    continue;

                    double angle = checkGridDirection.AngleTo(gridDirection);
#if Debug2020 || Revit2020
                    double angleDegrees = UnitUtils.ConvertFromInternalUnits(angle, DisplayUnitType.DUT_DECIMAL_DEGREES);
#else
                    double angleDegrees = UnitUtils.ConvertFromInternalUnits(angle, UnitTypeId.Degrees);
#endif
                    double normalizedAngle = angleDegrees > 90 ? 180 - angleDegrees : angleDegrees;
                    double roundedAngle = Math.Round(normalizedAngle, 5);
                    // Обнуляю разницу в 0,00001°, т.к. будет выдавать ПОЧТИ параллельные оси
                    if (tempNearestGrid == null || (roundedAngle > tempAngleDegrees && roundedAngle > 0.01 && roundedAngle < 89.99))
                    {
                        tempNearestAnglelGrid = grid;
                        tempAngleDegrees = roundedAngle;
                    }

                }
                // Иначе - проверяем расстояние
                else
                {
                    // Расстояние между параллельными линиями
                    IntersectionResult intRes1 = checkGridCurve.Project(fStartPontGrid);
                    IntersectionResult intRes2 = gridCurve.Project(fEndPontCheckGrid);
                    double distBetweenIntResPnts1 = intRes1.XYZPoint.DistanceTo(intRes2.XYZPoint);

                    IntersectionResult intRes3 = checkGridCurve.Project(sStartPontGrid);
                    IntersectionResult intRes4 = gridCurve.Project(sEndPontCheckGrid);
                    double distBetweenIntResPnts3 = intRes3.XYZPoint.DistanceTo(intRes4.XYZPoint);

                    double minBetweenDist = Math.Min(distBetweenIntResPnts1, distBetweenIntResPnts3);

                    // Проверяю угол между линией пересечений и проверяемой осью (если он не 0/90/180 - значит оси со смещением)
                    // Очень мелкие линии в ревит не построить. Просто игнор
                    if (minBetweenDist > 0.1)
                    {
                        Line lineBetween;
                        if (minBetweenDist == distBetweenIntResPnts1)
                            lineBetween = Line.CreateBound(intRes1.XYZPoint, intRes2.XYZPoint);
                        else
                            lineBetween = Line.CreateBound(intRes3.XYZPoint, intRes4.XYZPoint);

                        double angle = lineBetween.Direction.AngleTo(checkGridDirection);
#if Debug2020 || Revit2020
                        double angleDegrees = UnitUtils.ConvertFromInternalUnits(angle, DisplayUnitType.DUT_DECIMAL_DEGREES);
#else
                        double angleDegrees = UnitUtils.ConvertFromInternalUnits(angle, UnitTypeId.Degrees);
#endif
                        if (angleDegrees % 90 > 0.1)
                            minBetweenDist = minBetweenDist * Math.Sin(angle);
                    }

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
            }

            // Предварительно округляю с нужно точностью, чтобы и в отчет попало норм. значение, и анализировалось тоже корректное значение, а не напрм. 14.999999975
            double roundDist = Math.Round(tempDist, 5);

            // Заполняю данные по параллельным осям
            gridEnt.NearestParallelGrid = tempNearestGrid;
            gridEnt.NearestParallelDistance = roundDist;
            gridEnt.IsDistanceError = IsErrorValidNumber(roundDist);

            // Заполняю данные по НЕ параллельным осям
            //gridEnt.NearestAnglelGrid = tempNearestAnglelGrid;
            //gridEnt.NearestAngle = tempAngleDegrees;
            //gridEnt.IsAngleError = IsAngleError(tempAngleDegrees);

            // Помечаю, что элемент проверен
            gridEnt.IsChecked = true;
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

        private static bool IsAngleError(double angleDegrees)
        {
            double roundedAngle = Math.Round(angleDegrees, 0);
            double roundedToleranceAngle = Math.Round(angleDegrees, 5);

            if (Math.Abs(roundedAngle - roundedToleranceAngle) >= 0.000005)
                return true;

            return false;
        }
    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_MEPBender.Services.Insulation;
using KPLN_MEPBender.Services.Parameters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_MEPBender.Services.Routing
{
    public sealed class MepRouteBender
    {
        private readonly BendPathBuilder _bendPathBuilder;
        private readonly MepCurveFactory _mepCurveFactory;
        private readonly ObstacleOutlineBuilder _obstacleOutlineBuilder;
        private readonly InsulationSnapshotService _insulationSnapshotService;
        private readonly ParameterSnapshotService _parameterSnapshotService;

        public MepRouteBender()
        {
            _bendPathBuilder = new BendPathBuilder();
            _mepCurveFactory = new MepCurveFactory();
            _obstacleOutlineBuilder = new ObstacleOutlineBuilder();
            _parameterSnapshotService = new ParameterSnapshotService();
            _insulationSnapshotService = new InsulationSnapshotService(_parameterSnapshotService);
        }

        public MepBendResult Execute(MepBendRequest request)
        {
            MepBendResult result = new MepBendResult();

            using (Transaction transaction = new Transaction(request.Doc, "KPLN MEP Bender"))
            {
                transaction.Start();

                foreach (ElementId routeElementId in request.RouteElementIds)
                {
                    try
                    {
                        BendRoute(request, routeElementId, result);
                    }
                    catch (Exception ex)
                    {
                        result.FailedRouteCount++;
                        Element element = request.Doc.GetElement(routeElementId);
                        result.AddIssue(routeElementId, GetElementTypeName(element), "Ошибка трассы", GetExceptionText(ex));
                    }
                }

                transaction.Commit();
            }

            result.GeometryWasChanged = result.CreatedElementIds.Count > 0;
            result.Message = BuildResultMessage(result);
            return result;
        }

        private void BendRoute(MepBendRequest request, ElementId routeElementId, MepBendResult result)
        {
            MEPCurve source = request.Doc.GetElement(routeElementId) as MEPCurve;
            LocationCurve locationCurve = source?.Location as LocationCurve;
            Line routeLine = locationCurve?.Curve as Line;

            if (source == null || routeLine == null)
            {
                result.SkippedRouteCount++;
                result.AddIssue(routeElementId, GetElementTypeName(source), "Проверка трассы", source == null
                    ? "Элемент не найден или не является MEPCurve."
                    : "У элемента нет прямой LocationCurve. Сейчас поддерживаются только прямые участки.");
                return;
            }

            Outline obstacleOutline = _obstacleOutlineBuilder.BuildForRoute(request, source, 0, 0);
            if (obstacleOutline == null)
            {
                result.SkippedRouteCount++;
                result.AddIssue(routeElementId, GetElementTypeName(source), "Построение пути", "Для этой трассы нет пересекающихся препятствий. Непересекающиеся элементы выборки игнорируются.");
                return;
            }

            if (!_bendPathBuilder.TryBuild(request, source, routeLine, obstacleOutline, out List<XYZ> pathPoints))
            {
                result.SkippedRouteCount++;
                result.AddIssue(routeElementId, GetElementTypeName(source), "Построение пути", "Не удалось построить путь огибания: не хватает длины для угла/зазора или выбранное направление параллельно трассе.");
                return;
            }

            ParameterSnapshot sourceParameters = _parameterSnapshotService.Capture(source);
            List<InsulationSnapshot> insulationSnapshots = _insulationSnapshotService.Capture(request.Doc, source.Id);
            result.CapturedParameterSnapshotsCount++;

            XYZ sourceStart = routeLine.GetEndPoint(0);
            XYZ sourceEnd = routeLine.GetEndPoint(1);
            List<ExternalConnectorInfo> externalConnections = MepCurveConnectorUtils
                .GetExternalConnections(source, sourceStart, sourceEnd)
                .ToList();

            bool breakSucceeded = TryExecuteRouteSubTransaction(request.Doc, routeElementId, GetElementTypeName(source), "Разрыв исходника", routeResult =>
                TryBendByBreakingSource(request.Doc, source, pathPoints, sourceParameters, insulationSnapshots, routeResult), out MepBendResult breakResult);
            if (breakSucceeded)
            {
                MergeResult(result, breakResult);
                return;
            }

            source = request.Doc.GetElement(routeElementId) as MEPCurve;
            locationCurve = source?.Location as LocationCurve;
            MepBendResult keepResult = null;
            bool keepSucceeded = source != null && locationCurve != null && TryExecuteRouteSubTransaction(request.Doc, routeElementId, GetElementTypeName(source), "Запасное построение", routeResult =>
                TryBendByKeepingSourceAsFirstSegment(request.Doc, source, locationCurve, pathPoints, sourceParameters, insulationSnapshots, externalConnections, routeResult), out keepResult);
            if (keepSucceeded)
            {
                MergeResult(result, keepResult);
                return;
            }

            MergeIssues(result, breakResult);
            if (keepResult != null)
                MergeIssues(result, keepResult);

            result.SkippedRouteCount++;
            result.AddIssue(routeElementId, GetElementTypeName(source), "Итог трассы", "Не удалось выполнить ни разрыв исходного участка, ни запасное построение через укорочение исходника.");
        }

        private bool TryExecuteRouteSubTransaction(Document doc, ElementId routeElementId, string elementType, string stage, Func<MepBendResult, bool> routeAction, out MepBendResult routeResult)
        {
            routeResult = new MepBendResult();
            using (SubTransaction subTransaction = new SubTransaction(doc))
            {
                subTransaction.Start();
                try
                {
                    if (!routeAction(routeResult))
                    {
                        subTransaction.RollBack();
                        return false;
                    }

                    subTransaction.Commit();
                    return true;
                }
                catch (Exception ex)
                {
                    if (subTransaction.GetStatus() == TransactionStatus.Started)
                        subTransaction.RollBack();

                    routeResult.AddIssue(routeElementId, elementType, stage, GetExceptionText(ex));
                    return false;
                }
            }
        }

        private void MergeResult(MepBendResult result, MepBendResult routeResult)
        {
            result.CreatedElementIds.AddRange(routeResult.CreatedElementIds);
            result.ProcessedRouteIds.AddRange(routeResult.ProcessedRouteIds);
            result.CapturedParameterSnapshotsCount += routeResult.CapturedParameterSnapshotsCount;
            result.SkippedRouteCount += routeResult.SkippedRouteCount;
            result.FailedRouteCount += routeResult.FailedRouteCount;
            result.FittingFailureCount += routeResult.FittingFailureCount;
            result.ReconnectFailureCount += routeResult.ReconnectFailureCount;
            result.Issues.AddRange(routeResult.Issues);
        }

        private void MergeIssues(MepBendResult result, MepBendResult routeResult)
        {
            result.Issues.AddRange(routeResult.Issues);
        }

        private bool TryBendByBreakingSource(
            Document doc,
            MEPCurve source,
            List<XYZ> pathPoints,
            ParameterSnapshot sourceParameters,
            List<InsulationSnapshot> insulationSnapshots,
            MepBendResult result)
        {
            if (!(source is Pipe) && !(source is Duct))
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Разрыв исходника", "BreakCurve доступен только для труб и воздуховодов. Для этого типа будет использован запасной сценарий.");
                return false;
            }

            XYZ start = pathPoints.First();
            XYZ firstJoin = pathPoints[1];
            XYZ lastJoin = pathPoints[pathPoints.Count - 2];

            if (!TryBreakCurve(doc, source, firstJoin, out MEPCurve firstPieceCandidate, out MEPCurve secondPieceCandidate, out string breakError))
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Разрыв исходника", $"Не удалось разорвать исходный участок в первой точке огибания. {breakError}");
                return false;
            }

            MEPCurve headSegment = GetSegmentBetween(firstPieceCandidate, secondPieceCandidate, start, firstJoin);
            MEPCurve secondPart = headSegment != null && headSegment.Id.Equals(firstPieceCandidate.Id) ? secondPieceCandidate : firstPieceCandidate;

            if (headSegment == null || secondPart == null)
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Разрыв исходника", "После первого разрыва не удалось определить начальный и оставшийся фрагменты.");
                return false;
            }

            doc.Regenerate();

            if (!TryBreakCurve(doc, secondPart, lastJoin, out MEPCurve middleCandidate, out MEPCurve tailCandidate, out breakError))
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Разрыв исходника", $"Не удалось разорвать оставшийся участок во второй точке огибания. {breakError}");
                return false;
            }

            MEPCurve middleSegment = GetSegmentBetween(middleCandidate, tailCandidate, firstJoin, lastJoin);
            MEPCurve tailSegment = middleSegment != null && middleSegment.Id.Equals(middleCandidate.Id) ? tailCandidate : middleCandidate;

            if (middleSegment == null || tailSegment == null)
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Разрыв исходника", "После второго разрыва не удалось определить средний и конечный фрагменты.");
                return false;
            }

            _parameterSnapshotService.Apply(headSegment, sourceParameters);
            _parameterSnapshotService.Apply(tailSegment, sourceParameters);

            List<XYZ> bypassPoints = pathPoints.Skip(1).Take(pathPoints.Count - 2).ToList();
            List<MEPCurve> bypassSegments = CreateSegments(doc, source, bypassPoints, sourceParameters, insulationSnapshots, result);
            if (bypassSegments.Count != bypassPoints.Count - 1)
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Создание обхода", $"Создано участков: {bypassSegments.Count}, ожидалось: {bypassPoints.Count - 1}.");
                return false;
            }

            doc.Delete(middleSegment.Id);
            doc.Regenerate();

            List<MEPCurve> chainSegments = new List<MEPCurve> { headSegment };
            chainSegments.AddRange(bypassSegments);
            chainSegments.Add(tailSegment);

            List<XYZ> jointPoints = pathPoints.Skip(1).Take(pathPoints.Count - 2).ToList();
            result.FittingFailureCount += ConnectInternalSegments(doc, chainSegments, jointPoints, sourceParameters, result);

            result.ProcessedRouteIds.Add(source.Id);
            result.CreatedElementIds.AddRange(chainSegments.Where(s => !s.Id.Equals(source.Id)).Select(s => s.Id));
            return true;
        }

        private bool TryBendByKeepingSourceAsFirstSegment(
            Document doc,
            MEPCurve source,
            LocationCurve locationCurve,
            List<XYZ> pathPoints,
            ParameterSnapshot sourceParameters,
            List<InsulationSnapshot> insulationSnapshots,
            List<ExternalConnectorInfo> externalConnections,
            MepBendResult result)
        {
            try
            {
                locationCurve.Curve = Line.CreateBound(pathPoints[0], pathPoints[1]);
            }
            catch (Exception ex)
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Запасное построение", $"Не удалось укоротить исходный участок: {GetExceptionText(ex)}");
                return false;
            }

            List<MEPCurve> createdSegments = CreateSegments(doc, source, pathPoints.Skip(1).ToList(), sourceParameters, insulationSnapshots, result);
            if (createdSegments.Count < 1)
            {
                result.AddIssue(source.Id, GetElementTypeName(source), "Запасное построение", "Не удалось создать новые участки после укорочения исходника.");
                return false;
            }

            List<MEPCurve> chainSegments = new List<MEPCurve> { source };
            chainSegments.AddRange(createdSegments);

            doc.Regenerate();
            List<XYZ> jointPoints = pathPoints.Skip(1).Take(pathPoints.Count - 2).ToList();
            result.FittingFailureCount += ConnectInternalSegments(doc, chainSegments, jointPoints, sourceParameters, result);
            doc.Regenerate();

            result.ReconnectFailureCount += ReconnectExternalElements(doc, chainSegments, pathPoints, externalConnections, result);

            result.ProcessedRouteIds.Add(source.Id);
            result.CreatedElementIds.AddRange(createdSegments.Select(s => s.Id));
            return true;
        }

        private List<MEPCurve> CreateSegments(
            Document doc,
            MEPCurve source,
            List<XYZ> pathPoints,
            ParameterSnapshot sourceParameters,
            List<InsulationSnapshot> insulationSnapshots,
            MepBendResult result)
        {
            List<MEPCurve> newSegments = new List<MEPCurve>();

            for (int i = 1; i < pathPoints.Count; i++)
            {
                MEPCurve segment = null;
                try
                {
                    segment = _mepCurveFactory.Create(doc, source, pathPoints[i - 1], pathPoints[i]);
                }
                catch (Exception ex)
                {
                    result.AddIssue(source.Id, GetElementTypeName(source), "Создание участка", $"Участок {i}: {GetExceptionText(ex)}");
                    continue;
                }

                if (segment == null)
                {
                    result.AddIssue(source.Id, GetElementTypeName(source), "Создание участка", $"Участок {i}: Revit не создал элемент или длина участка меньше минимума.");
                    continue;
                }

                try
                {
                    _parameterSnapshotService.Apply(segment, sourceParameters);
                    result.CreatedElementIds.AddRange(_insulationSnapshotService.Apply(doc, segment.Id, insulationSnapshots));
                }
                catch (Exception ex)
                {
                    result.AddIssue(segment.Id, GetElementTypeName(segment), "Перенос параметров", GetExceptionText(ex));
                }

                newSegments.Add(segment);
            }

            return newSegments;
        }

        private int ConnectInternalSegments(Document doc, List<MEPCurve> newSegments, List<XYZ> jointPoints, ParameterSnapshot sourceParameters, MepBendResult result)
        {
            int failureCount = 0;

            for (int i = 1; i < newSegments.Count; i++)
            {
                XYZ jointPoint = jointPoints[i - 1];
                Connector firstConnector = MepCurveConnectorUtils.GetClosestEndConnector(newSegments[i - 1], jointPoint);
                Connector secondConnector = MepCurveConnectorUtils.GetClosestEndConnector(newSegments[i], jointPoint);

                if (!MepCurveConnectorUtils.TryConnect(doc, firstConnector, secondConnector, true, out Element fitting))
                {
                    failureCount++;
                    result.AddIssue(newSegments[i].Id, GetElementTypeName(newSegments[i]), "Соединение", $"Не удалось соединить участки в точке {FormatPoint(jointPoint)}.");
                    continue;
                }

                if (fitting != null)
                {
                    try
                    {
                        _parameterSnapshotService.Apply(fitting, sourceParameters);
                    }
                    catch (Exception ex)
                    {
                        result.AddIssue(fitting.Id, GetElementTypeName(fitting), "Параметры фитинга", GetExceptionText(ex));
                    }

                    result.CreatedElementIds.Add(fitting.Id);
                }
            }

            return failureCount;
        }

        private bool TryBreakCurve(Document doc, MEPCurve curve, XYZ point, out MEPCurve firstCandidate, out MEPCurve secondCandidate, out string error)
        {
            firstCandidate = null;
            secondCandidate = null;
            error = string.Empty;

            ElementId newElementId = ElementId.InvalidElementId;

            try
            {
                Pipe pipe = curve as Pipe;
                if (pipe != null)
                    newElementId = PlumbingUtils.BreakCurve(doc, pipe.Id, point);

                Duct duct = curve as Duct;
                if (duct != null)
                    newElementId = MechanicalUtils.BreakCurve(doc, duct.Id, point);
            }
            catch (Exception ex)
            {
                error = GetExceptionText(ex);
                return false;
            }

            if (newElementId == ElementId.InvalidElementId)
            {
                error = "BreakCurve вернул InvalidElementId.";
                return false;
            }

            doc.Regenerate();

            firstCandidate = doc.GetElement(curve.Id) as MEPCurve;
            secondCandidate = doc.GetElement(newElementId) as MEPCurve;

            if (firstCandidate == null || secondCandidate == null)
            {
                error = "После BreakCurve не найден один из созданных фрагментов.";
                return false;
            }

            return true;
        }

        private bool IsSegmentBetween(MEPCurve curve, XYZ firstPoint, XYZ secondPoint)
        {
            LocationCurve locationCurve = curve?.Location as LocationCurve;
            Line line = locationCurve?.Curve as Line;
            if (line == null)
                return false;

            XYZ start = line.GetEndPoint(0);
            XYZ end = line.GetEndPoint(1);
            double direct = start.DistanceTo(firstPoint) + end.DistanceTo(secondPoint);
            double reverse = start.DistanceTo(secondPoint) + end.DistanceTo(firstPoint);

            return Math.Min(direct, reverse) < 0.01;
        }

        private MEPCurve GetSegmentBetween(MEPCurve firstCandidate, MEPCurve secondCandidate, XYZ firstPoint, XYZ secondPoint)
        {
            if (IsSegmentBetween(firstCandidate, firstPoint, secondPoint))
                return firstCandidate;

            if (IsSegmentBetween(secondCandidate, firstPoint, secondPoint))
                return secondCandidate;

            return null;
        }

        private int ReconnectExternalElements(
            Document doc,
            List<MEPCurve> newSegments,
            List<XYZ> pathPoints,
            List<ExternalConnectorInfo> externalConnections,
            MepBendResult result)
        {
            int failureCount = 0;
            MEPCurve firstSegment = newSegments.First();
            MEPCurve lastSegment = newSegments.Last();

            foreach (ExternalConnectorInfo externalConnection in externalConnections)
            {
                Element owner = doc.GetElement(externalConnection.OwnerId);
                if (owner == null)
                    continue;

                bool reconnectToStart = externalConnection.SourceEndpoint.DistanceTo(pathPoints.First())
                                        <= externalConnection.SourceEndpoint.DistanceTo(pathPoints.Last());

                MEPCurve targetSegment = reconnectToStart ? firstSegment : lastSegment;
                XYZ endpoint = reconnectToStart ? pathPoints.First() : pathPoints.Last();

                Connector newConnector = MepCurveConnectorUtils.GetClosestEndConnector(targetSegment, endpoint);
                Connector externalConnector = MepCurveConnectorUtils.GetClosestConnector(owner, externalConnection.SourceEndpoint);

                if (!MepCurveConnectorUtils.TryConnect(doc, newConnector, externalConnector, false, out Element createdElement))
                {
                    failureCount++;
                    result.AddIssue(targetSegment.Id, GetElementTypeName(targetSegment), "Переподключение", $"Не удалось переподключить внешний элемент id {externalConnection.OwnerId}.");
                }
            }

            return failureCount;
        }

        private string BuildResultMessage(MepBendResult result)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append($"Обработано трасс: {result.ProcessedRouteIds.Count}. Создано участков: {result.CreatedElementIds.Count}. Пропущено: {result.SkippedRouteCount}. Ошибок трасс: {result.FailedRouteCount}. Ошибок фитингов: {result.FittingFailureCount}. Ошибок переподключения: {result.ReconnectFailureCount}.");

            if (result.Issues.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("Детали:");
                foreach (RouteIssue issue in result.Issues.Take(12))
                    builder.AppendLine(issue.ToString());

                if (result.Issues.Count > 12)
                    builder.AppendLine($"... и ещё {result.Issues.Count - 12} сообщений.");
            }

            return builder.ToString();
        }

        private string GetElementTypeName(Element element)
        {
            if (element == null)
                return "Unknown";

            if (element is Pipe)
                return "Pipe";

            if (element is Duct)
                return "Duct";

            if (element is CableTray)
                return "CableTray";

            return element.GetType().Name;
        }

        private string GetExceptionText(Exception ex)
        {
            if (ex == null)
                return "Неизвестная ошибка.";

            return string.IsNullOrWhiteSpace(ex.Message)
                ? ex.GetType().Name
                : $"{ex.GetType().Name}: {ex.Message}";
        }

        private string FormatPoint(XYZ point)
        {
            if (point == null)
                return "-";

            return $"{point.X:0.###}; {point.Y:0.###}; {point.Z:0.###}";
        }
    }
}

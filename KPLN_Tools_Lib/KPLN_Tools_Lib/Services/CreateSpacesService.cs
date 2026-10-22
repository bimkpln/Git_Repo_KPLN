using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_Lib.Services
{
    internal sealed class CreateSpacesService
    {
        // Revit хранит длины в футах. Допуск сопоставления отметок — 1 мм.
#if Debug2020 || Revit2020
        private const DisplayUnitType Millimeters = DisplayUnitType.DUT_MILLIMETERS;
#else
        private static readonly ForgeTypeId Millimeters = UnitTypeId.Millimeters;
#endif

        internal static readonly double PositionTolerance = UnitUtils.ConvertToInternalUnits(1.0, Millimeters);

        internal CreateSpacesResult Create(Document document, IEnumerable<RevitLinkInstance> links, Phase phase)
        {
            var result = new CreateSpacesResult();
            Level[] levels = new FilteredElementCollector(document).OfClass(typeof(Level)).Cast<Level>().ToArray();
            List<Space> spaces = new FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
                .Cast<Space>().Where(space => GetPhaseId(space).Equals(phase.Id)).ToList();

            // Как в исходном Python: одна транзакция на все выбранные связи и помещения.
            using (var transaction = new Transaction(document, "KPLN: Создать пространства по помещениям"))
            {
                transaction.Start();
                var failures = new SpaceFailures(result.FailureIdsByMessage);
                FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(failures);
                options.SetClearAfterRollback(true);
                transaction.SetFailureHandlingOptions(options);

                foreach (RevitLinkInstance link in links)
                {
                    Document linkedDocument = link.GetLinkDocument();
                    if (linkedDocument == null)
                    {
                        result.Messages.Add(string.Format("Связь «{0}» выгружена и пропущена.", link.Name));
                        continue;
                    }

                    Transform transform = link.GetTransform();
                    if (!transform.BasisZ.IsAlmostEqualTo(XYZ.BasisZ))
                    {
                        result.Messages.Add(string.Format("Связь «{0}» наклонена: сопоставление горизонтальных уровней невозможно.", link.Name));
                        continue;
                    }

                    // Берём помещения стадии, сопоставленной со стадией пространств в настройках связи.
                    ElementId linkedPhaseId;
                    try
                    {
                        var linkType = (RevitLinkType)document.GetElement(link.GetTypeId());
                        if (!linkType.GetPhaseMap().TryGetValue(phase.Id, out linkedPhaseId))
                            throw new InvalidOperationException("Не задано соответствие стадий.");
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                    {
                        result.Messages.Add(string.Format("Связь «{0}»: {1}", link.Name, ex.Message));
                        continue;
                    }
                    catch (InvalidOperationException ex)
                    {
                        result.Messages.Add(string.Format("Связь «{0}»: {1}", link.Name, ex.Message));
                        continue;
                    }

                    Room[] rooms = new FilteredElementCollector(linkedDocument)
                        .OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
                        .Cast<Room>().Where(room => GetPhaseId(room).Equals(linkedPhaseId)).ToArray();

                    if (rooms.Length == 0)
                        result.Messages.Add(string.Format("Связь «{0}»: нет помещений на сопоставленной стадии.", link.Name));

                    foreach (Room room in rooms)
                        CreateForRoom(document, link, transform, room, levels, spaces, phase, result);
                }

                if (result.Created > 0)
                {
                    TransactionStatus status = transaction.Commit();
                    result.Messages.AddRange(failures.Messages);
                    if (status != TransactionStatus.Committed)
                        throw new InvalidOperationException("Revit не подтвердил сохранение созданных пространств.");
                }
                else
                    transaction.RollBack();
            }

            return result;
        }

        private static void CreateForRoom(Document document, RevitLinkInstance link, Transform transform,
            Room room, Level[] levels, List<Space> spaces, Phase phase, CreateSpacesResult result)
        {
            string roomLabel = string.Format("Связь «{0}» [id: {1}], помещение «{2} {3}» [id: {4}]",
                link.Name, link.Id, room.Number, room.Name, room.Id);

            try
            {
                if (!(room.Location is LocationPoint location) || room.Area <= 0)
                {
                    Skip(result, roomLabel, "помещение не размещено, не замкнуто или имеет нулевую площадь.");
                    return;
                }

                Level roomLevel = room.Level;
                if (roomLevel == null)
                {
                    Skip(result, roomLabel, "не определён уровень помещения.");
                    return;
                }

                XYZ point = transform.OfPoint(location.Point);
                double elevation = transform.OfPoint(new XYZ(0, 0, roomLevel.ProjectElevation)).Z;
                Level[] matches = levels.Where(level => Math.Abs(level.ProjectElevation - elevation) <= PositionTolerance).ToArray();
                Level hostLevel = matches.Length == 1 ? matches[0] : null;

                if (matches.Length > 1)
                {
                    Level[] namesakes = matches.Where(level => level.Name == roomLevel.Name).ToArray();
                    if (namesakes.Length == 1)
                        hostLevel = namesakes[0];
                }

                if (hostLevel == null)
                {
                    Skip(result, roomLabel, string.Format(
                        "для уровня «{0}» (отметка в текущем проекте {1:F1} мм) {2}.",
                        roomLevel.Name, UnitUtils.ConvertFromInternalUnits(elevation, Millimeters),
                        matches.Length == 0 ? "нет соответствующего уровня" : "найдено несколько уровней"));
                    return;
                }

                double height = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble();
                if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0)
                {
                    Skip(result, roomLabel, "некорректная высота помещения.");
                    return;
                }

                double baseOffset = elevation + room.BaseOffset - hostLevel.ProjectElevation;
                XYZ insidePoint = new XYZ(point.X, point.Y, hostLevel.ProjectElevation + baseOffset + height / 2);
                bool sameLocation = spaces.Any(existingSpace => existingSpace.LevelId.Equals(hostLevel.Id)
                    && existingSpace.Location is LocationPoint existing
                    && Math.Abs(existing.Point.X - point.X) <= PositionTolerance
                    && Math.Abs(existing.Point.Y - point.Y) <= PositionTolerance);

                if (sameLocation || document.GetSpaceAtPoint(insidePoint, phase) != null)
                {
                    result.Existing++;
                    return;
                }


                Space space = document.Create.NewSpace(hostLevel, phase, new UV(point.X, point.Y))
                    ?? throw new InvalidOperationException("Revit не создал пространство.");

                // Как в исходном Python: высота помещения записывается в верхнее смещение.
                try
                {
                    if (!space.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(height))
                        throw new InvalidOperationException("Revit не записал верхнее смещение пространства.");
                }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException)
                {
                    throw;
                }
                catch
                {
                    document.Delete(space.Id);
                    throw;
                }

                spaces.Add(space);
                result.Created++;
            }
            catch (Autodesk.Revit.Exceptions.RegenerationFailedException)
            {
                // После ошибки регенерации документ нельзя читать: откатываем общую транзакцию.
                throw;
            }
            catch (Exception ex)
            {
                Skip(result, roomLabel, ex.Message);
            }
        }

        private static ElementId GetPhaseId(Element element)
            => element.get_Parameter(BuiltInParameter.ROOM_PHASE_ID)?.AsElementId() ?? element.CreatedPhaseId;

        private static void Skip(CreateSpacesResult result, string label, string reason)
        {
            result.Skipped++;
            result.Messages.Add(label + ": " + reason);
        }

        private sealed class SpaceFailures : IFailuresPreprocessor
        {
            internal List<string> Messages { get; } = new List<string>();

            private readonly Dictionary<string, List<ElementId>> _failureIdsByMessage;

            internal SpaceFailures(Dictionary<string, List<ElementId>> failureIdsByMessage)
            {
                _failureIdsByMessage = failureIdsByMessage;
            }

            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                bool hasErrors = false;
                foreach (FailureMessageAccessor failure in accessor.GetFailureMessages())
                {
                    // Одно сообщение — одна коллекция ID текущей модели без повторов.
                    string description = failure.GetDescriptionText();
                    ICollection<ElementId> failingIds = failure.GetFailingElementIds();
                    if (failingIds.Count > 0)
                    {
                        if (!_failureIdsByMessage.TryGetValue(description, out List<ElementId> ids))
                        {
                            ids = new List<ElementId>();
                            _failureIdsByMessage.Add(description, ids);
                        }

                        foreach (ElementId id in failingIds)
                        {
                            if (!ids.Contains(id))
                                ids.Add(id);
                        }
                    }
                    else
                    {
                        string message = description + " — Revit не сообщил ID элементов.";
                        if (!Messages.Contains(message))
                            Messages.Add(message);
                    }

                    if (failure.GetSeverity() == FailureSeverity.Warning)
                    {
                        // Предупреждения сохраняются в итоговом отчёте, без серии модальных окон.
                        accessor.DeleteWarning(failure);
                    }
                    else
                        hasErrors = true;
                }

                return hasErrors ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}

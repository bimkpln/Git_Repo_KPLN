using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_Furniture3DFrom2D : IExternalCommand
    {
        private const string FurnitureLibrary =
            @"X:\BIM\3_Семейства\1_АР\000_Архитектурная концепция\140_Мебель (3D)";

        private const string PlumbingLibrary =
            @"X:\BIM\3_Семейства\1_АР\000_Архитектурная концепция\700_Сантехнические приборы (3D)";

        private static readonly BuiltInCategory[] Categories =
        {
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_PlumbingFixtures
        };

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                ConversionDirection? direction = ShowDirectionDialog();
                if (!direction.HasValue)
                    return Result.Cancelled;

                IList<Reference> references;
                try
                {
                    references = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new FurnitureAndPlumbingFilter(),
                        "Выберите элементы мебели и сантехники на виде");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                // Общее вложенное семейство является подкомпонентом и управляется
                // родительским экземпляром. При выборе такого элемента заменяем
                // верхний размещённый экземпляр семейства.
                List<FamilyInstance> pickedElements = references
                    .Select(x => doc.GetElement(x) as FamilyInstance)
                    .Where(x => x != null)
                    .Select(GetTopLevelFamilyInstance)
                    .GroupBy(x => x.Id.IntegerValue)
                    .Select(x => x.First())
                    .ToList();

                if (pickedElements.Count == 0)
                {
                    TaskDialog.Show("Внимание!", "Действие отменено. Замена семейств не выполнена.");
                    return Result.Cancelled;
                }

                Dictionary<ElementId, List<ParameterValue>> savedParameters =
                    pickedElements.ToDictionary(x => x.Id, ReadWritableParameters);
                Dictionary<ElementId, List<NestedInstanceSnapshot>> savedNestedParameters =
                    pickedElements.ToDictionary(x => x.Id, x => ReadNestedParameters(doc, x));

                if (direction.Value == ConversionDirection.To3D)
                    LoadMissing3DFamilies(doc, pickedElements);

                Dictionary<int, FurnitureReplacementResult> resultBySourceId =
                    pickedElements.ToDictionary(
                        x => x.Id.IntegerValue,
                        x => new FurnitureReplacementResult(
                            x.Id.IntegerValue,
                            x.Symbol.FamilyName + " : " + GetTypeName(x.Symbol)));
                Dictionary<int, ElementId> replacementIds = new Dictionary<int, ElementId>();

                using (Transaction transaction = new Transaction(
                    doc,
                    direction.Value == ConversionDirection.To3D
                        ? "Заменить 2D семейства на 3D"
                        : "Заменить 3D семейства на 2D"))
                {
                    transaction.Start();

                    List<FamilySymbol> candidates = GetCandidateSymbols(doc, direction.Value);
                    List<GroupSnapshot> affectedGroups = UngroupAffectedGroups(
                        doc,
                        pickedElements,
                        resultBySourceId);

                    foreach (FamilyInstance source in pickedElements)
                    {
                        FurnitureReplacementResult row = resultBySourceId[source.Id.IntegerValue];
                        if (row.Status == FurnitureReplacementStatus.Failed)
                            continue;

                        string sourceFamilyName = source.Symbol.FamilyName;

                        // Уже подходящие экземпляры заменять не требуется.
                        if (IsAlreadyInTargetForm(sourceFamilyName, direction.Value))
                        {
                            row.MarkSuccess("Замена не требовалась");
                            continue;
                        }

                        LocationPoint sourceLocation = source.Location as LocationPoint;
                        if (sourceLocation == null)
                        {
                            row.MarkFailed("Поддерживаются только семейства с точечным размещением.");
                            continue;
                        }

                        string sourceTypeName = GetTypeName(source.Symbol);
                        MatchResult match = FindMatchingSymbol(
                            candidates,
                            sourceFamilyName,
                            sourceTypeName,
                            direction.Value);

                        if (match.Symbol == null)
                        {
                            row.MarkFailed("Семейство отсутствует в " +
                                (direction.Value == ConversionDirection.To3D ? "3D" : "2D") +
                                " библиотеке.");
                            continue;
                        }

                        if (!match.Symbol.IsActive)
                        {
                            match.Symbol.Activate();
                            doc.Regenerate();
                        }

                        PlacementSnapshot placement = new PlacementSnapshot(
                            sourceLocation.Point,
                            sourceLocation.Rotation,
                            source.HandOrientation,
                            source.FacingOrientation,
                            source.HandFlipped,
                            source.FacingFlipped,
                            source.Mirrored,
                            ReadPlanFootprint(source, doc.ActiveView));
                        ElementId sourceId = source.Id;

                        using (SubTransaction itemTransaction = new SubTransaction(doc))
                        {
                            itemTransaction.Start();
                            try
                            {
                                doc.Delete(sourceId);
                                FamilyInstance created = doc.Create.NewFamilyInstance(
                                    placement.Point,
                                    match.Symbol,
                                    StructuralType.NonStructural);

                                // В исходном скрипте запасной типоразмер получал
                                // только тот же Location.Rotation. Не добавляем к
                                // нему неподтверждённые поправки на 90° или 180°.
                                if (!match.ExactTypeMatch)
                                {
                                    RotateAtPoint(
                                        doc,
                                        created,
                                        placement.Point,
                                        placement.Rotation);
                                }

                                RestoreParameters(created, savedParameters[sourceId]);
                                doc.Regenerate();
                                RestoreNestedParameters(
                                    doc,
                                    created,
                                    savedNestedParameters[sourceId]);
                                doc.Regenerate();

                                bool fallbackMirroringRestored = true;
                                if (match.ExactTypeMatch)
                                {
                                    RestorePlacement(doc, created, placement);
                                }
                                else
                                {
                                    RestoreLocationPoint(doc, created, placement.Point);
                                    fallbackMirroringRestored = RestoreFallbackMirroring(
                                        doc,
                                        created,
                                        placement);
                                }

                                replacementIds[sourceId.IntegerValue] = created.Id;
                                row.Id = created.Id.IntegerValue;

                                if (!match.ExactTypeMatch)
                                {
                                    string warning = "Исходный типоразмер \"" + sourceTypeName +
                                        "\" не найден; использован \"" +
                                        GetTypeName(match.Symbol) + "\".";

                                    warning += GetFallbackPlacementDiagnostics(
                                        created,
                                        placement,
                                        doc.ActiveView);

                                    if (!fallbackMirroringRestored)
                                    {
                                        warning += " Исходное отражение восстановить " +
                                            "не удалось.";
                                    }

                                    row.MarkWarning(warning);
                                }
                                else
                                {
                                    row.MarkSuccess(string.Empty);
                                }

                                itemTransaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                itemTransaction.RollBack();
                                replacementIds.Remove(sourceId.IntegerValue);
                                row.Id = sourceId.IntegerValue;
                                row.MarkFailed(ex.Message);
                            }
                        }
                    }

                    RestoreGroups(
                        doc,
                        affectedGroups,
                        replacementIds,
                        resultBySourceId);

                    transaction.Commit();
                }

                ShowResultWindow(
                    resultBySourceId.Values);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private static ConversionDirection? ShowDirectionDialog()
        {
            TaskDialog dialog = new TaskDialog("Заменить семейства")
            {
                MainInstruction = "Выберите направление замены:",
                MainContent = "После этого выберите на виде экземпляры семейств мебели и сантехники.",
                CommonButtons = TaskDialogCommonButtons.Cancel,
                DefaultButton = TaskDialogResult.Cancel,
                AllowCancellation = true
            };

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Заменить 2D на 3D");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Заменить 3D на 2D");

            TaskDialogResult result = dialog.Show();
            if (result == TaskDialogResult.CommandLink1)
                return ConversionDirection.To3D;
            if (result == TaskDialogResult.CommandLink2)
                return ConversionDirection.To2D;
            return null;
        }

        private static void LoadMissing3DFamilies(
            Document doc,
            IEnumerable<FamilyInstance> instances)
        {
            IEnumerable<string> familyPaths = instances
                .Select(x => x.Symbol.FamilyName)
                .Where(x => !x.EndsWith("_3d", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(name =>
                {
                    if (name.StartsWith("140_", StringComparison.OrdinalIgnoreCase))
                        return Path.Combine(FurnitureLibrary, name + "_3d.rfa");
                    if (name.StartsWith("700_", StringComparison.OrdinalIgnoreCase))
                        return Path.Combine(PlumbingLibrary, name + "_3d.rfa");
                    return null;
                })
                .Where(path => !string.IsNullOrEmpty(path) && File.Exists(path));

            using (Transaction transaction = new Transaction(doc, "Загрузить семейства"))
            {
                transaction.Start();
                foreach (string path in familyPaths)
                    doc.LoadFamily(path);
                transaction.Commit();
            }
        }

        private static List<FamilySymbol> GetCandidateSymbols(
            Document doc,
            ConversionDirection direction)
        {
            List<FamilySymbol> result = new List<FamilySymbol>();

            foreach (BuiltInCategory category in Categories)
            {
                IEnumerable<FamilySymbol> symbols = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsElementType()
                    .OfType<FamilySymbol>();

                result.AddRange(symbols.Where(x =>
                    direction == ConversionDirection.To3D
                        ? x.FamilyName.EndsWith("_3d", StringComparison.OrdinalIgnoreCase)
                        : !x.FamilyName.EndsWith("_3d", StringComparison.OrdinalIgnoreCase)));
            }

            return result;
        }

        private static MatchResult FindMatchingSymbol(
            IEnumerable<FamilySymbol> candidates,
            string sourceFamilyName,
            string sourceTypeName,
            ConversionDirection direction)
        {
            string targetFamilyName = direction == ConversionDirection.To3D
                ? sourceFamilyName + "_3d"
                : Remove3DSuffix(sourceFamilyName);

            List<FamilySymbol> familySymbols = candidates
                .Where(x => string.Equals(
                    x.FamilyName,
                    targetFamilyName,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (familySymbols.Count == 0)
                return new MatchResult(null, false);

            string alternativeTypeName = direction == ConversionDirection.To3D
                ? sourceTypeName + "_3d"
                : Remove3DSuffix(sourceTypeName);

            FamilySymbol exact = familySymbols.FirstOrDefault(x =>
                string.Equals(GetTypeName(x), sourceTypeName, StringComparison.OrdinalIgnoreCase));

            if (exact == null)
            {
                exact = familySymbols.FirstOrDefault(x =>
                    string.Equals(GetTypeName(x), alternativeTypeName, StringComparison.OrdinalIgnoreCase));
            }

            return exact != null
                ? new MatchResult(exact, true)
                : new MatchResult(familySymbols[0], false);
        }

        private static List<ParameterValue> ReadWritableParameters(FamilyInstance instance)
        {
            List<ParameterValue> result = new List<ParameterValue>();

            foreach (Parameter parameter in instance.Parameters)
            {
                if (parameter == null || parameter.IsReadOnly || !parameter.HasValue ||
                    parameter.Definition == null)
                    continue;

                if (parameter.StorageType == StorageType.Double)
                {
                    result.Add(new ParameterValue(
                        parameter.Definition.Name,
                        StorageType.Double,
                        parameter.AsDouble(),
                        0));
                }
                else if (parameter.StorageType == StorageType.Integer)
                {
                    result.Add(new ParameterValue(
                        parameter.Definition.Name,
                        StorageType.Integer,
                        0.0,
                        parameter.AsInteger()));
                }
            }

            return result;
        }

        private static void RestoreParameters(
            FamilyInstance instance,
            IEnumerable<ParameterValue> values)
        {
            foreach (ParameterValue value in values)
            {
                Parameter parameter = instance.LookupParameter(value.Name);
                if (parameter == null || parameter.IsReadOnly ||
                    parameter.StorageType != value.StorageType)
                    continue;

                try
                {
                    if (value.StorageType == StorageType.Double)
                        parameter.Set(value.DoubleValue);
                    else if (value.StorageType == StorageType.Integer)
                        parameter.Set(value.IntegerValue);
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    // Формулы и часть встроенных параметров могут не принимать значение.
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Значение может быть недопустимо для параметра нового семейства.
                }
            }
        }

        private static List<NestedInstanceSnapshot> ReadNestedParameters(
            Document doc,
            FamilyInstance root)
        {
            List<NestedInstanceSnapshot> result = new List<NestedInstanceSnapshot>();

            foreach (FamilyInstance nested in GetNestedFamilyInstances(doc, root))
            {
                result.Add(new NestedInstanceSnapshot(
                    GetMatchingKey(nested),
                    ReadWritableParameters(nested)));
            }

            return result;
        }

        private static void RestoreNestedParameters(
            Document doc,
            FamilyInstance root,
            IEnumerable<NestedInstanceSnapshot> snapshots)
        {
            Dictionary<string, Queue<NestedInstanceSnapshot>> byKey = snapshots
                .GroupBy(x => x.MatchingKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    x => x.Key,
                    x => new Queue<NestedInstanceSnapshot>(x),
                    StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance nested in GetNestedFamilyInstances(doc, root))
            {
                string key = GetMatchingKey(nested);
                Queue<NestedInstanceSnapshot> queue;

                if (!byKey.TryGetValue(key, out queue) || queue.Count == 0)
                    continue;

                RestoreParameters(nested, queue.Dequeue().Parameters);
            }
        }

        private static IEnumerable<FamilyInstance> GetNestedFamilyInstances(
            Document doc,
            FamilyInstance root)
        {
            Queue<FamilyInstance> pending = new Queue<FamilyInstance>();
            HashSet<int> visited = new HashSet<int>();
            pending.Enqueue(root);

            while (pending.Count > 0)
            {
                FamilyInstance parent = pending.Dequeue();

                foreach (ElementId id in parent.GetSubComponentIds())
                {
                    if (!visited.Add(id.IntegerValue))
                        continue;

                    FamilyInstance nested = doc.GetElement(id) as FamilyInstance;
                    if (nested == null)
                        continue;

                    yield return nested;
                    pending.Enqueue(nested);
                }
            }
        }

        private static string GetMatchingKey(FamilyInstance instance)
        {
            return Remove3DSuffix(instance.Symbol.FamilyName) + "\n" +
                   Remove3DSuffix(GetTypeName(instance.Symbol));
        }

        private static void RotateAtPoint(
            Document doc,
            FamilyInstance instance,
            XYZ point,
            double angle)
        {
            if (Math.Abs(angle) < 1e-9)
                return;

            Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
            ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle);
        }

        private static void RestorePlacement(
            Document doc,
            FamilyInstance instance,
            PlacementSnapshot placement)
        {
            RestoreFlipState(instance, placement);
            doc.Regenerate();

            LocationPoint location = instance.Location as LocationPoint;
            if (location == null)
                throw new InvalidOperationException(
                    "Созданное семейство не поддерживает точечное размещение.");

            XYZ currentDirection = GetPlanDirection(instance.HandOrientation);
            XYZ targetDirection = GetPlanDirection(placement.HandOrientation);

            if (currentDirection == null || targetDirection == null)
            {
                currentDirection = GetPlanDirection(instance.FacingOrientation);
                targetDirection = GetPlanDirection(placement.FacingOrientation);
            }

            double rotation = currentDirection != null && targetDirection != null
                ? GetSignedPlanAngle(currentDirection, targetDirection)
                : placement.Rotation - location.Rotation;

            RotateAtPoint(doc, instance, location.Point, rotation);
            doc.Regenerate();

            location = instance.Location as LocationPoint;
            if (location == null)
                throw new InvalidOperationException(
                    "Не удалось определить точку размещения созданного семейства.");

            XYZ translation = placement.Point - location.Point;
            if (translation.GetLength() > 1e-9)
                ElementTransformUtils.MoveElement(doc, instance.Id, translation);
        }

        private static void RestoreLocationPoint(
            Document doc,
            FamilyInstance instance,
            XYZ targetPoint)
        {
            LocationPoint location = instance.Location as LocationPoint;
            if (location == null)
                throw new InvalidOperationException(
                    "Не удалось определить точку размещения созданного семейства.");

            XYZ translation = targetPoint - location.Point;
            if (translation.GetLength() > 1e-9)
                ElementTransformUtils.MoveElement(doc, instance.Id, translation);
        }

        private static bool RestoreFallbackMirroring(
            Document doc,
            FamilyInstance instance,
            PlacementSnapshot placement)
        {
            if (instance.Mirrored == placement.Mirrored)
                return true;

            if (!ElementTransformUtils.CanMirrorElement(doc, instance.Id))
                return false;

            LocationPoint location = instance.Location as LocationPoint;
            if (location == null)
                return false;

            bool facingMustChange =
                instance.FacingFlipped != placement.FacingFlipped;
            bool handMustChange =
                instance.HandFlipped != placement.HandFlipped;

            // Плоскость с нормалью вдоль FacingOrientation отражает лицевую
            // ось, сохраняя HandOrientation. Для HandFlipped — наоборот.
            XYZ mirrorNormal;
            if (facingMustChange && !handMustChange)
                mirrorNormal = GetPlanDirection(instance.FacingOrientation);
            else if (handMustChange && !facingMustChange)
                mirrorNormal = GetPlanDirection(instance.HandOrientation);
            else
                mirrorNormal = GetPlanDirection(instance.FacingOrientation);

            if (mirrorNormal == null)
                return false;

            Plane mirrorPlane = Plane.CreateByNormalAndOrigin(
                mirrorNormal,
                location.Point);

            ElementTransformUtils.MirrorElements(
                doc,
                new List<ElementId> { instance.Id },
                mirrorPlane,
                false);
            doc.Regenerate();
            RestoreLocationPoint(doc, instance, placement.Point);

            return instance.Mirrored == placement.Mirrored;
        }

        private static PlanFootprint ReadPlanFootprint(
            FamilyInstance instance,
            View view)
        {
            // Берём именно видимый на текущем плане отпечаток: у 2D-семейств
            // символические линии могут отсутствовать в модельном bounding box.
            BoundingBoxXYZ bounds = instance.get_BoundingBox(view);
            if (bounds == null)
                return null;

            Transform transform = bounds.Transform;
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;

            for (int x = 0; x < 2; x++)
            {
                for (int y = 0; y < 2; y++)
                {
                    for (int z = 0; z < 2; z++)
                    {
                        XYZ corner = new XYZ(
                            x == 0 ? bounds.Min.X : bounds.Max.X,
                            y == 0 ? bounds.Min.Y : bounds.Max.Y,
                            z == 0 ? bounds.Min.Z : bounds.Max.Z);
                        XYZ worldCorner = transform.OfPoint(corner);

                        minX = Math.Min(minX, worldCorner.X);
                        minY = Math.Min(minY, worldCorner.Y);
                        maxX = Math.Max(maxX, worldCorner.X);
                        maxY = Math.Max(maxY, worldCorner.Y);
                    }
                }
            }

            return new PlanFootprint(
                new XYZ((minX + maxX) * 0.5, (minY + maxY) * 0.5, 0.0),
                maxX - minX,
                maxY - minY);
        }

        private static string GetFallbackPlacementDiagnostics(
            FamilyInstance instance,
            PlacementSnapshot placement,
            View view)
        {
            LocationPoint location = instance.Location as LocationPoint;
            double resultingRotation = location != null
                ? location.Rotation
                : double.NaN;

            string text = " Диагностика положения: угол " +
                FormatAngleDegrees(placement.Rotation) + " -> " +
                FormatAngleDegrees(resultingRotation) +
                "; отражение " + FormatBoolean(placement.Mirrored) +
                " -> " + FormatBoolean(instance.Mirrored) +
                "; H/F " + FormatBoolean(placement.HandFlipped) + "/" +
                FormatBoolean(placement.FacingFlipped) + " -> " +
                FormatBoolean(instance.HandFlipped) + "/" +
                FormatBoolean(instance.FacingFlipped) + ".";

            PlanFootprint createdFootprint = ReadPlanFootprint(instance, view);
            if (placement.Footprint != null && createdFootprint != null)
            {
                text += " Отпечаток X/Y: " +
                    FormatLengthMillimeters(placement.Footprint.SizeX) + "x" +
                    FormatLengthMillimeters(placement.Footprint.SizeY) + " -> " +
                    FormatLengthMillimeters(createdFootprint.SizeX) + "x" +
                    FormatLengthMillimeters(createdFootprint.SizeY) + " мм.";
            }

            return text;
        }

        private static string FormatAngleDegrees(double angle)
        {
            if (double.IsNaN(angle))
                return "не определён";

            double degrees = angle * 180.0 / Math.PI;
            degrees %= 360.0;
            if (degrees < 0.0)
                degrees += 360.0;

            return degrees.ToString("0.#") + "°";
        }

        private static string FormatBoolean(bool value)
        {
            return value ? "да" : "нет";
        }

        private static string FormatLengthMillimeters(double internalLength)
        {
            return Math.Round(internalLength * 304.8).ToString("0");
        }

        private static void RestoreFlipState(
            FamilyInstance instance,
            PlacementSnapshot placement)
        {
            // У разных семейств начальное состояние разворота может отличаться.
            // После каждой операции повторно проверяем оба признака: у некоторых
            // семейств один flip одновременно изменяет HandFlipped и FacingFlipped.
            for (int attempt = 0; attempt < 2; attempt++)
            {
                if (instance.HandFlipped != placement.HandFlipped &&
                    instance.CanFlipHand)
                {
                    instance.flipHand();
                }

                if (instance.FacingFlipped != placement.FacingFlipped &&
                    instance.CanFlipFacing)
                {
                    instance.flipFacing();
                }

                if (instance.HandFlipped == placement.HandFlipped &&
                    instance.FacingFlipped == placement.FacingFlipped)
                {
                    break;
                }
            }
        }

        private static XYZ GetPlanDirection(XYZ direction)
        {
            if (direction == null)
                return null;

            XYZ projected = new XYZ(direction.X, direction.Y, 0.0);
            return projected.GetLength() > 1e-9 ? projected.Normalize() : null;
        }

        private static double GetSignedPlanAngle(XYZ from, XYZ to)
        {
            double dot = from.X * to.X + from.Y * to.Y;
            double crossZ = from.X * to.Y - from.Y * to.X;
            return Math.Atan2(crossZ, dot);
        }

        private static string GetTypeName(FamilySymbol symbol)
        {
            Parameter parameter = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
            return parameter != null ? parameter.AsString() ?? string.Empty : string.Empty;
        }

        private static string Remove3DSuffix(string value)
        {
            return value != null && value.EndsWith("_3d", StringComparison.OrdinalIgnoreCase)
                ? value.Substring(0, value.Length - 3)
                : value;
        }

        private static bool IsAlreadyInTargetForm(
            string familyName,
            ConversionDirection direction)
        {
            bool is3D = familyName.EndsWith("_3d", StringComparison.OrdinalIgnoreCase);
            return direction == ConversionDirection.To3D ? is3D : !is3D;
        }

        private static FamilyInstance GetTopLevelFamilyInstance(FamilyInstance instance)
        {
            FamilyInstance current = instance;

            while (current.SuperComponent is FamilyInstance)
                current = (FamilyInstance)current.SuperComponent;

            return current;
        }

        private static List<GroupSnapshot> UngroupAffectedGroups(
            Document doc,
            IEnumerable<FamilyInstance> instances,
            IDictionary<int, FurnitureReplacementResult> results)
        {
            List<GroupSnapshot> snapshots = new List<GroupSnapshot>();
            IEnumerable<ElementId> groupIds = instances
                .Where(x => x.GroupId != ElementId.InvalidElementId)
                .Select(x => x.GroupId)
                .GroupBy(x => x.IntegerValue)
                .Select(x => x.First());

            foreach (ElementId groupId in groupIds)
            {
                Group group = doc.GetElement(groupId) as Group;
                if (group == null)
                    continue;

                List<ElementId> members = group.GetMemberIds().ToList();
                bool hasOtherInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .Cast<Group>()
                    .Count(x => x.GetTypeId() == group.GetTypeId()) > 1;

                GroupSnapshot snapshot = new GroupSnapshot(
                    group.Name,
                    members,
                    hasOtherInstances);

                try
                {
                    group.UngroupMembers();
                    snapshots.Add(snapshot);
                }
                catch (Exception ex)
                {
                    foreach (FamilyInstance selected in instances.Where(x => x.GroupId == groupId))
                    {
                        FurnitureReplacementResult row;
                        if (results.TryGetValue(selected.Id.IntegerValue, out row))
                            row.MarkFailed("Не удалось расформировать группу: " + ex.Message);
                    }
                }
            }

            return snapshots;
        }

        private static void RestoreGroups(
            Document doc,
            IEnumerable<GroupSnapshot> snapshots,
            IDictionary<int, ElementId> replacementIds,
            IDictionary<int, FurnitureReplacementResult> results)
        {
            foreach (GroupSnapshot snapshot in snapshots)
            {
                List<ElementId> currentMembers = snapshot.MemberIds
                    .Select(x => replacementIds.ContainsKey(x.IntegerValue)
                        ? replacementIds[x.IntegerValue]
                        : x)
                    .Where(x => doc.GetElement(x) != null)
                    .ToList();

                try
                {
                    if (currentMembers.Count == 0)
                        throw new InvalidOperationException("В группе не осталось доступных элементов.");

                    doc.Create.NewGroup(currentMembers);

                    if (snapshot.HadOtherInstances)
                    {
                        AddGroupWarning(
                            snapshot,
                            results,
                            "Группа восстановлена, но получила отдельный тип; " +
                            "исходный тип имел другие экземпляры.");
                    }
                }
                catch (Exception ex)
                {
                    AddGroupWarning(
                        snapshot,
                        results,
                        "Не удалось собрать группу \"" + snapshot.Name + "\" обратно: " + ex.Message);
                }
            }
        }

        private static void AddGroupWarning(
            GroupSnapshot snapshot,
            IDictionary<int, FurnitureReplacementResult> results,
            string text)
        {
            foreach (ElementId memberId in snapshot.MemberIds)
            {
                FurnitureReplacementResult row;
                if (results.TryGetValue(memberId.IntegerValue, out row) &&
                    row.Status != FurnitureReplacementStatus.Failed)
                    row.MarkWarning(text);
            }
        }

        private static void ShowResultWindow(
            IEnumerable<FurnitureReplacementResult> results)
        {
            SelectFurnitureElementHandler selectionHandler =
                new SelectFurnitureElementHandler();
            ExternalEvent selectionEvent = ExternalEvent.Create(selectionHandler);

            Furniture2DFrom3D window = new Furniture2DFrom3D(
                selectionEvent: selectionEvent,
                selectionHandler: selectionHandler,
                results: results);
            window.Show();
        }

        private sealed class FurnitureAndPlumbingFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                FamilyInstance instance = element as FamilyInstance;
                if (instance == null || instance.Category == null)
                    return false;

                int categoryId = instance.Category.Id.IntegerValue;
                bool allowedCategory =
                    categoryId == (int)BuiltInCategory.OST_Furniture ||
                    categoryId == (int)BuiltInCategory.OST_PlumbingFixtures;

                if (!allowedCategory)
                    return false;

                // Общие семейства, включая общие вложенные компоненты,
                // также должны быть доступны для выбора. Если выбран вложенный
                // компонент, Execute поднимется до его родительского экземпляра.
                return true;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        private enum ConversionDirection
        {
            To3D,
            To2D
        }

        private sealed class MatchResult
        {
            public MatchResult(FamilySymbol symbol, bool exactTypeMatch)
            {
                Symbol = symbol;
                ExactTypeMatch = exactTypeMatch;
            }

            public FamilySymbol Symbol { get; private set; }
            public bool ExactTypeMatch { get; private set; }
        }

        private sealed class ParameterValue
        {
            public ParameterValue(
                string name,
                StorageType storageType,
                double doubleValue,
                int integerValue)
            {
                Name = name;
                StorageType = storageType;
                DoubleValue = doubleValue;
                IntegerValue = integerValue;
            }

            public string Name { get; private set; }
            public StorageType StorageType { get; private set; }
            public double DoubleValue { get; private set; }
            public int IntegerValue { get; private set; }
        }

        private sealed class PlacementSnapshot
        {
            public PlacementSnapshot(
                XYZ point,
                double rotation,
                XYZ handOrientation,
                XYZ facingOrientation,
                bool handFlipped,
                bool facingFlipped,
                bool mirrored,
                PlanFootprint footprint)
            {
                Point = point;
                Rotation = rotation;
                HandOrientation = handOrientation;
                FacingOrientation = facingOrientation;
                HandFlipped = handFlipped;
                FacingFlipped = facingFlipped;
                Mirrored = mirrored;
                Footprint = footprint;
            }

            public XYZ Point { get; private set; }
            public double Rotation { get; private set; }
            public XYZ HandOrientation { get; private set; }
            public XYZ FacingOrientation { get; private set; }
            public bool HandFlipped { get; private set; }
            public bool FacingFlipped { get; private set; }
            public bool Mirrored { get; private set; }
            public PlanFootprint Footprint { get; private set; }
        }

        private sealed class PlanFootprint
        {
            public PlanFootprint(XYZ center, double sizeX, double sizeY)
            {
                Center = center;
                SizeX = sizeX;
                SizeY = sizeY;
            }

            public XYZ Center { get; private set; }
            public double SizeX { get; private set; }
            public double SizeY { get; private set; }
        }

        private sealed class NestedInstanceSnapshot
        {
            public NestedInstanceSnapshot(
                string matchingKey,
                List<ParameterValue> parameters)
            {
                MatchingKey = matchingKey;
                Parameters = parameters;
            }

            public string MatchingKey { get; private set; }
            public List<ParameterValue> Parameters { get; private set; }
        }

        private sealed class GroupSnapshot
        {
            public GroupSnapshot(
                string name,
                List<ElementId> memberIds,
                bool hadOtherInstances)
            {
                Name = name;
                MemberIds = memberIds;
                HadOtherInstances = hadOtherInstances;
            }

            public string Name { get; private set; }
            public List<ElementId> MemberIds { get; private set; }
            public bool HadOtherInstances { get; private set; }
        }
    }
}
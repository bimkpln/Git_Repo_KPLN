using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace KPLN_Tools.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler : IExternalEventHandler
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string ApartmentInstanceMarker = "[KPLN_APT_INSTANCE]";

        private static readonly Dictionary<string, string> _familyNameByPathCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private ApartmentManagerWindow _window;

        private static void ApplyApartmentFailureHandling(Transaction transaction)
        {
            if (transaction == null)
                return;

            FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
            if (options == null)
                return;

            options.SetFailuresPreprocessor(new ApartmentWarningSuppressor());
            transaction.SetFailureHandlingOptions(options);
        }

        private class ApartmentWarningSuppressor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                if (failuresAccessor == null)
                    return FailureProcessingResult.Continue;

                IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
                if (failures == null || failures.Count == 0)
                    return FailureProcessingResult.Continue;

                foreach (FailureMessageAccessor failure in failures.ToList())
                {
                    if (failure == null || failure.GetSeverity() != FailureSeverity.Warning)
                        continue;

                    string description = null;
                    try
                    {
                        description = failure.GetDescriptionText();
                    }
                    catch
                    {
                        description = null;
                    }

                    if (ShouldSuppressApartmentWarning(description))
                        failuresAccessor.DeleteWarning(failure);
                }

                return FailureProcessingResult.Continue;
            }
        }

        private static bool ShouldSuppressApartmentWarning(string description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return false;

            string normalized = description
                .Replace('ё', 'е')
                .ToLowerInvariant();

            return normalized.Contains("слегка отклони") ||
                   normalized.Contains("slightly off axis");
        }

        public string GetName()
        {
            return "KPLN. Менеджер квартир";
        }

        public void AttachWindow(ApartmentManagerWindow window)
        {
            _window = window;
        }

        public void DetachWindow()
        {
            _window = null;
        }

        private void RestoreWindow()
        {
            if (_window == null)
                return;

            if (_window.Dispatcher == null)
                return;

            _window.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_window == null)
                    return;

                if (!_window.IsVisible)
                    return;

                if (_window.WindowState == WindowState.Minimized)
                    _window.WindowState = WindowState.Normal;

                _window.Show();
                _window.Activate();

                _window.Topmost = true;
                _window.Topmost = false;
                _window.Focus();
            }), DispatcherPriority.Background);
        }

        private class ApartmentFamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = false;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = false;
                return true;
            }
        }

        private class ExistingWallLineInfo
        {
            public ElementId WallId { get; set; }
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
            public double ThicknessInternal { get; set; }
        }

        private class PreparedApartmentWalls
        {
            public ElementId ApartmentId { get; set; }
            public WallType WallType { get; set; }
            public int ThicknessMm { get; set; }
            public List<Line> AxisLines { get; set; }
            public WallType ShaftWallType { get; set; }
            public List<Line> ShaftAxisLines { get; set; }
            public WallType LoggiaWallType { get; set; }
            public List<Line> LoggiaAxisLines { get; set; }
            public List<Line> RoomSeparatorLines { get; set; }
        }

        private class ApartmentProcessState
        {
            public ElementId ApartmentId { get; set; }
            public string ApartmentFamilyName { get; set; }
            public ElementId NavigationElementId { get; set; }
            public List<ElementId> NavigationElementIds { get; set; }
            public List<ElementId> CreatedElementIds { get; set; }
            public Apartment2DRestoreInfo Restore2DInfo { get; set; }
            public bool HasPreparedWalls { get; set; }
            public bool HasCreatedWalls { get; set; }
            public bool HasCreatedRooms { get; set; }
            public bool HasInstalledDoors { get; set; }
            public bool HasInstalledWindows { get; set; }
            public int FoundEntranceDoorsCount { get; set; }
            public int InstalledEntranceDoorsCount { get; set; }
            public int FoundRoomSeparatorsCount { get; set; }
            public int CreatedRoomSeparatorsCount { get; set; }
            public int SkippedRoomsCount { get; set; }
            public int SkippedWallsCount { get; set; }
            public int SkippedDoorsCount { get; set; }
            public int SkippedWindowsCount { get; set; }
            public List<string> FurnitureErrors { get; set; }
            public List<string> ErrorMessages { get; set; }
            public bool HasRoomAreaMismatch { get; set; }

            public ApartmentProcessState()
            {
                NavigationElementId = ElementId.InvalidElementId;
                NavigationElementIds = new List<ElementId>();
                CreatedElementIds = new List<ElementId>();
                FurnitureErrors = new List<string>();
                ErrorMessages = new List<string>();
            }
        }

        private class OffsetLine2D
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        private class Interval1D
        {
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisLineData
        {
            public XYZ Dir { get; set; }
            public XYZ Normal { get; set; }
            public double Offset { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisGroupKey : IEquatable<GenericAxisGroupKey>
        {
            public XYZ Dir { get; private set; }
            public double Offset { get; private set; }
            public double Z { get; private set; }

            public GenericAxisGroupKey(XYZ dir, double offset, double z, double tol)
            {
                Dir = new XYZ(RoundTol(dir.X, tol), RoundTol(dir.Y, tol), 0);
                Offset = RoundTol(offset, tol);
                Z = RoundTol(z, tol);
            }

            public bool Equals(GenericAxisGroupKey other)
            {
                if (other == null)
                    return false;

                return
                    Math.Abs(Dir.X - other.Dir.X) < 1e-9 &&
                    Math.Abs(Dir.Y - other.Dir.Y) < 1e-9 &&
                    Math.Abs(Offset - other.Offset) < 1e-9 &&
                    Math.Abs(Z - other.Z) < 1e-9;
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as GenericAxisGroupKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + Dir.X.GetHashCode();
                    hash = hash * 23 + Dir.Y.GetHashCode();
                    hash = hash * 23 + Offset.GetHashCode();
                    hash = hash * 23 + Z.GetHashCode();
                    return hash;
                }
            }
        }

        private class FamilyWindowMarker
        {
            public XYZ LocalP0 { get; set; }
            public XYZ LocalP1 { get; set; }
        }

        private class FamilyShaftWallMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class FamilyLoggiaWallMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class FamilyRoomSeparatorMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class HelperLineCandidate
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public bool StyleMatched { get; set; }
        }

        private class DoorTypeMirrorEnsureResult
        {
            public bool HasMessage { get; set; }
            public string Message { get; set; }
        }

        private enum DoorOpeningMarker
        {
            None,
            Left,
            Right,
            RightAlt
        }

        private class PreparedDoorPlacement
        {
            public ElementId ApartmentId { get; set; }
            public ElementId Door2DId { get; set; }
            public string RoomCategory { get; set; }
            public int DoorWidthMm { get; set; }
            public string SelectedDoorTypeName { get; set; }
            public FamilySymbol DoorSymbol { get; set; }
            public XYZ InsertPoint { get; set; }
            public FamilyInstance RelatedRoom2D { get; set; }
            public XYZ InteriorReferencePoint { get; set; }
            public XYZ SourceHandDirection { get; set; }
            public XYZ SourceFacingDirection { get; set; }
            public bool IsEntranceDoor { get; set; }
            public List<string> Diagnostics { get; set; }
            public bool RequiresOppositeDoorTypeAfterWallFlip { get; set; }
        }

        private class PreparedApartmentDoors
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedDoorPlacement> Doors { get; set; }

            public PreparedApartmentDoors()
            {
                Doors = new List<PreparedDoorPlacement>();
            }
        }

        private class PreparedWindowPlacement
        {
            public ElementId ApartmentId { get; set; }
            public FamilySymbol WindowSymbol { get; set; }
            public Line SourceLine { get; set; }
            public XYZ InsertPoint { get; set; }
            public XYZ ReferenceDirection { get; set; }
            public double SillHeightInternal { get; set; }
            public List<string> Diagnostics { get; set; }
        }

        private class PreparedApartmentWindows
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedWindowPlacement> Windows { get; set; }

            public PreparedApartmentWindows()
            {
                Windows = new List<PreparedWindowPlacement>();
            }
        }

        private class PreparedRoomPlacement
        {
            public ElementId ApartmentId { get; set; }
            public string RoomName { get; set; }
            public XYZ InsertPoint { get; set; }
            public double ExpectedAreaInternal { get; set; }
            public bool HasShaftInside { get; set; }
        }

        private class PreparedApartmentRooms
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedRoomPlacement> Rooms { get; set; }
            public bool HasRoomSeparators { get; set; }
            public bool HasShafts { get; set; }

            public PreparedApartmentRooms()
            {
                Rooms = new List<PreparedRoomPlacement>();
            }
        }

        private class RoomAreaMismatchInfo
        {
            public ElementId ApartmentId { get; set; }
            public string RoomName { get; set; }
            public double ExpectedAreaInternal { get; set; }
            public double ActualAreaInternal { get; set; }
        }

        private enum RequestType
        {
            None,
            PlaceApartment,
            ConvertTo3D,
            RefreshApartmentPresets,
            UpdateApartmentMarks
        }

        private RequestType _requestType = RequestType.None;
        private int _requestedApartmentId;
        private ApartmentPresetData _requestedPresetData;

        public void Execute(UIApplication app)
        {
            try
            {
                UIDocument uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                    throw new Exception("Не найден активный UIDocument.");

                Document doc = uidoc.Document;
                if (doc == null)
                    throw new Exception("Не найден активный документ.");

                switch (_requestType)
                {
                    case RequestType.PlaceApartment:
                        if (ExecutePlaceApartment(uidoc, doc, _requestedApartmentId))
                            MarkApartmentPresetDataStaleInWindow();
                        break;

                    case RequestType.ConvertTo3D:
                        string validationMessage;
                        if (!ValidatePresetBeforeConvertTo3D(doc, _requestedPresetData, out validationMessage))
                        {
                            TaskDialog.Show("Предупреждение", validationMessage);
                            return;
                        }

                        ExecuteConvertTo3D(uidoc, doc, _requestedPresetData);
                        break;

                    case RequestType.RefreshApartmentPresets:
                        ExecuteRefreshApartmentPresets(doc, _requestedPresetData);
                        break;

                    case RequestType.UpdateApartmentMarks:
                        ExecuteUpdateApartmentMarks(doc);
                        break;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
            }
            finally
            {
                _requestType = RequestType.None;
                _requestedApartmentId = 0;
                _requestedPresetData = null;

                RestoreWindow();
            }
        }
        public void PreparePlaceApartment(int apartmentId)
        {
            _requestType = RequestType.PlaceApartment;
            _requestedApartmentId = apartmentId;
            _requestedPresetData = null;
        }

        public void PrepareConvertTo3D(ApartmentPresetData presetData)
        {
            _requestType = RequestType.ConvertTo3D;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        public void PrepareRefreshApartmentPresets(ApartmentPresetData presetData)
        {
            _requestType = RequestType.RefreshApartmentPresets;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        public void PrepareUpdateApartmentMarks()
        {
            _requestType = RequestType.UpdateApartmentMarks;
            _requestedApartmentId = 0;
            _requestedPresetData = null;
        }

        private void MarkApartmentPresetDataStaleInWindow()
        {
            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                _window.MarkApartmentPresetDataStale();
            }));
        }

        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? ("Data Source=" + dbPath + ";Version=3;Read Only=True;")
                : ("Data Source=" + dbPath + ";Version=3;");

            SQLiteConnection con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

        private static string GetFamilyPathById(int id)
        {
            if (!File.Exists(DbPath))
                throw new FileNotFoundException("Не найдена база данных", DbPath);

            using (SQLiteConnection con = OpenConnection(DbPath, true))
            using (SQLiteCommand cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT FPATH FROM Main WHERE ID = @id LIMIT 1;";
                cmd.Parameters.AddWithValue("@id", id);

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    return null;

                return result.ToString().Trim();
            }
        }

        private static string GetFamilyNameFromFile(Document projectDoc, string familyPath)
        {
            if (projectDoc == null)
                return null;

            if (string.IsNullOrWhiteSpace(familyPath) || !File.Exists(familyPath))
                return null;

            string cachedName;
            if (_familyNameByPathCache.TryGetValue(familyPath, out cachedName))
                return string.IsNullOrWhiteSpace(cachedName) ? null : cachedName;

            Document familyDoc = null;
            string detectedName = null;

            try
            {
                familyDoc = projectDoc.Application.OpenDocumentFile(familyPath);

                Family family = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault();

                if (family != null && !string.IsNullOrWhiteSpace(family.Name))
                    detectedName = family.Name.Trim();
            }
            catch
            {
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }

            _familyNameByPathCache[familyPath] = detectedName ?? "";
            return detectedName;
        }

        private static double RoundTol(double value, double tol)
        {
            return Math.Round(value / tol) * tol;
        }

        private static ElementId GetInstanceLevelId(FamilyInstance fi)
        {
            if (fi == null)
                return ElementId.InvalidElementId;

            Parameter p =
                fi.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                fi.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);

            if (p != null && p.StorageType == StorageType.ElementId)
                return p.AsElementId();

            if (fi.LevelId != ElementId.InvalidElementId)
                return fi.LevelId;

            return ElementId.InvalidElementId;
        }

        private static string GetCommentsValue(Element e)
        {
            if (e == null)
                return null;

            List<string> instanceValues = GetCommentParameterValues(e, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            string entranceInstanceValue = instanceValues.FirstOrDefault(IsEntranceDoorComment);
            if (!string.IsNullOrWhiteSpace(entranceInstanceValue))
                return entranceInstanceValue.Trim();

            string firstInstanceValue = instanceValues.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
            if (!string.IsNullOrWhiteSpace(firstInstanceValue))
                return firstInstanceValue.Trim();

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                List<string> typeValues = GetCommentParameterValues(typeElem, BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
                string entranceTypeValue = typeValues.FirstOrDefault(IsEntranceDoorComment);
                if (!string.IsNullOrWhiteSpace(entranceTypeValue))
                    return entranceTypeValue.Trim();

                string firstTypeValue = typeValues.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                if (!string.IsNullOrWhiteSpace(firstTypeValue))
                    return firstTypeValue.Trim();
            }

            return null;
        }

        private static bool HasEntranceDoorComment(Element e)
        {
            if (e == null)
                return false;

            List<string> instanceValues = GetCommentParameterValues(e, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            return instanceValues.Any(IsEntranceDoorComment);
        }

        private static List<string> GetCommentParameterValues(Element e, BuiltInParameter builtInParameter)
        {
            List<string> values = new List<string>();
            if (e == null)
                return values;

            AddParameterValue(values, e.get_Parameter(builtInParameter));
            AddNamedParameterValues(values, e, "Комментарии");
            AddNamedParameterValues(values, e, "Комментарий");
            AddNamedParameterValues(values, e, "Comments");
            AddNamedParameterValues(values, e, "Comment");

            return values;
        }

        private static void AddNamedParameterValues(List<string> values, Element e, string parameterName)
        {
            if (values == null || e == null || string.IsNullOrWhiteSpace(parameterName))
                return;

            try
            {
                IList<Parameter> parameters = e.GetParameters(parameterName);
                if (parameters != null)
                {
                    foreach (Parameter parameter in parameters)
                        AddParameterValue(values, parameter);
                }
            }
            catch
            {
            }

            try
            {
                AddParameterValue(values, e.LookupParameter(parameterName));
            }
            catch
            {
            }
        }

        private static void AddParameterValue(List<string> values, Parameter parameter)
        {
            if (values == null || parameter == null)
                return;

            string value = null;

            try
            {
                if (parameter.StorageType == StorageType.String)
                    value = parameter.AsString();
                else
                    value = parameter.AsValueString();
            }
            catch
            {
                value = null;
            }

            if (string.IsNullOrWhiteSpace(value))
                return;

            value = value.Trim();
            if (!values.Contains(value))
                values.Add(value);
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesInDocument(Document doc)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            if (doc == null)
                return result;

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (FamilyInstance fi in instances)
            {
                if (fi == null)
                    continue;

                Parameter pComment = fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pComment == null)
                    continue;

                string comment = pComment.AsString();
                if (string.IsNullOrWhiteSpace(comment))
                    continue;

                if (!comment.Contains(ApartmentInstanceMarker))
                    continue;

                result.Add(fi);
            }

            return result;
        }

        private static bool SetApartmentAreaParameterValue(FamilyInstance apartmentFi, string parameterName, double areaInternal, List<string> errors)
        {
            if (apartmentFi == null || string.IsNullOrWhiteSpace(parameterName))
                return false;

            Parameter p = apartmentFi.LookupParameter(parameterName);
            if (p == null)
            {
                if (errors != null)
                    errors.Add("У квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " не найден параметр '" + parameterName + "'.");
                return false;
            }

            if (p.IsReadOnly)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " доступен только для чтения.");
                return false;
            }

            if (p.StorageType != StorageType.Double)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " имеет некорректный тип.");
                return false;
            }

            p.Set(areaInternal);
            return true;
        }

        private void ExecuteUpdateApartmentMarks(Document doc)
        {
            List<FamilyInstance> apartments = GetPlacedApartmentInstancesInDocument(doc);

            if (apartments.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "В модели не найдены квартиры, размещённые через менеджер.");
                return;
            }

            int updatedCount = 0;
            int skippedCount = 0;
            List<string> errors = new List<string>();

            using (Transaction t = new Transaction(doc, "KPLN. Обновление марок квартир"))
            {
                t.Start();

                foreach (FamilyInstance apartmentFi in apartments)
                {
                    if (apartmentFi == null)
                        continue;

                    try
                    {
                        List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);

                        double livingAreaInternal = 0.0;
                        double totalAreaInternal = 0.0;

                        foreach (FamilyInstance roomFi in roomInstances)
                        {
                            if (roomFi == null)
                                continue;

                            double roomAreaInternal;
                            if (!TryGetAreaParamFromElementOrType(roomFi, out roomAreaInternal, "КП_Р_Площадь", "КП_Р_ПЛОЩАДЬ"))
                                continue;

                            totalAreaInternal += roomAreaInternal;

                            string roomCategory = GetRoomCategoryLabel(roomFi);
                            if (string.Equals((roomCategory ?? "").Trim(), "Комната", StringComparison.OrdinalIgnoreCase))
                                livingAreaInternal += roomAreaInternal;
                        }

                        bool livingOk = SetApartmentAreaParameterValue(apartmentFi, "КВ_Площадь_Жилая", livingAreaInternal, errors);
                        bool totalOk = SetApartmentAreaParameterValue(apartmentFi, "КВ_Площадь_Общая", totalAreaInternal, errors);

                        if (livingOk && totalOk)
                            updatedCount++;
                        else
                            skippedCount++;
                    }
                    catch (Exception ex)
                    {
                        skippedCount++;
                        errors.Add("Ошибка у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + ": " + ex.Message);
                    }
                }

                t.Commit();
            }

            string message =
                "Найдено квартир: " + apartments.Count +
                "\nОбновлено: " + updatedCount +
                "\nПропущено: " + skippedCount;

            if (errors.Count > 0)
            {
                List<string> shortErrors = errors.Take(15).ToList();
                message += "\n\nОшибки:\n- " + string.Join("\n- ", shortErrors);

                if (errors.Count > shortErrors.Count)
                    message += "\n- ...";
            }

            TaskDialog.Show("KPLN. Менеджер квартир", message);
        }

        private bool ExecutePlaceApartment(UIDocument uidoc, Document doc, int id)
        {
            ViewPlan floorPlan = doc.ActiveView as ViewPlan;
            if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("Предупреждение", "Откройте план этажа перед размещением квартиры.");
                return false;
            }

            string familyPath = GetFamilyPathById(id);
            if (string.IsNullOrWhiteSpace(familyPath))
            {
                TaskDialog.Show("Ошибка", "Для ID = " + id + " не найден FPATH в базе.");
                return false;
            }

            if (!File.Exists(familyPath))
            {
                TaskDialog.Show("Ошибка", "Файл семейства не найден:\n" + familyPath);
                return false;
            }

            FamilySymbol symbol = null;

            using (Transaction t = new Transaction(doc, "Загрузка семейства квартиры"))
            {
                t.Start();

                Family family = LoadOrFindFamily(doc, familyPath);
                if (family == null)
                    throw new Exception("Не удалось загрузить или найти семейство в проекте.");

                symbol = GetFirstFamilySymbol(doc, family);
                if (symbol == null)
                    throw new Exception("В семействе не найден ни один типоразмер.");

                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }

                t.Commit();
            }

            int placedCount = 0;

            using (TransactionGroup tg = new TransactionGroup(doc, "Размещение семейства квартиры"))
            {
                tg.Start();

                while (true)
                {
                    XYZ insertPoint;

                    try
                    {
                        insertPoint = uidoc.Selection.PickPoint("Укажите точку вставки квартиры. ESC - завершить.");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        break;
                    }

                    using (Transaction t = new Transaction(doc, "Размещение семейства квартиры"))
                    {
                        t.Start();

                        FamilyInstance placedInstance = PlaceFamilyInstance(doc, floorPlan, symbol, insertPoint);
                        if (placedInstance == null)
                            throw new Exception("Не удалось разместить семейство квартиры.");

                        AppendComment(placedInstance, ApartmentInstanceMarker);

                        t.Commit();
                    }

                    placedCount++;
                }

                if (placedCount > 0)
                    tg.Assimilate();
                else
                    tg.RollBack();
            }

            return placedCount > 0;
        }

        private static Family LoadOrFindFamily(Document doc, string familyPath)
        {
            if (doc == null)
                throw new ArgumentNullException("doc");

            if (string.IsNullOrWhiteSpace(familyPath))
                throw new ArgumentException("Не задан путь к семейству.", "familyPath");

            if (!File.Exists(familyPath))
                throw new FileNotFoundException("Файл семейства не найден.", familyPath);

            string fileFamilyName = Path.GetFileNameWithoutExtension(familyPath);

            List<Family> existingFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            Family existingByFileName = existingFamilies.FirstOrDefault(f =>
                string.Equals(f.Name, fileFamilyName, StringComparison.OrdinalIgnoreCase));

            if (existingByFileName != null)
                return existingByFileName;

            string realFamilyName = GetFamilyNameFromFile(doc, familyPath);

            if (!string.IsNullOrWhiteSpace(realFamilyName))
            {
                Family existingByRealName = existingFamilies.FirstOrDefault(f =>
                    string.Equals(f.Name, realFamilyName, StringComparison.OrdinalIgnoreCase));

                if (existingByRealName != null)
                    return existingByRealName;
            }

            Family loadedFamily = null;

            try
            {
                bool loaded = doc.LoadFamily(familyPath, new ApartmentFamilyLoadOptions(), out loadedFamily);

                if (loaded && loadedFamily != null)
                    return loadedFamily;
            }
            catch
            {
            }

            existingFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            if (!string.IsNullOrWhiteSpace(realFamilyName))
            {
                Family existingAfterLoadByRealName = existingFamilies.FirstOrDefault(f =>
                    string.Equals(f.Name, realFamilyName, StringComparison.OrdinalIgnoreCase));

                if (existingAfterLoadByRealName != null)
                    return existingAfterLoadByRealName;
            }

            Family existingAfterLoadByFileName = existingFamilies.FirstOrDefault(f =>
                string.Equals(f.Name, fileFamilyName, StringComparison.OrdinalIgnoreCase));

            if (existingAfterLoadByFileName != null)
                return existingAfterLoadByFileName;

            return null;
        }

        private static FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
            if (symbolIds == null || symbolIds.Count == 0)
                return null;

            ElementId firstId = symbolIds.First();
            return doc.GetElement(firstId) as FamilySymbol;
        }

        private static FamilyInstance PlaceFamilyInstance(Document doc, ViewPlan floorPlan, FamilySymbol symbol, XYZ point)
        {
            FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;

            switch (placementType)
            {
                case FamilyPlacementType.ViewBased:
                    return doc.Create.NewFamilyInstance(point, symbol, floorPlan);

                case FamilyPlacementType.OneLevelBased:
                case FamilyPlacementType.OneLevelBasedHosted:
                    if (floorPlan.GenLevel == null)
                        throw new Exception("У активного плана не определён уровень.");

                    return doc.Create.NewFamilyInstance(point, symbol, floorPlan.GenLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                case FamilyPlacementType.WorkPlaneBased:
                    if (floorPlan.GenLevel == null)
                        throw new Exception("У активного плана не определён уровень.");

                    return doc.Create.NewFamilyInstance(point, symbol, floorPlan.GenLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                default:
                    throw new NotSupportedException("Тип размещения семейства не поддерживается: " + placementType);
            }
        }

        private static void AppendComment(Element e, string textToAppend)
        {
            if (e == null || string.IsNullOrWhiteSpace(textToAppend))
                return;

            Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p == null || p.IsReadOnly)
                return;

            string oldValue = p.AsString();

            if (string.IsNullOrWhiteSpace(oldValue))
            {
                p.Set(textToAppend);
                return;
            }

            if (oldValue.Contains(textToAppend))
                return;

            p.Set(oldValue + " " + textToAppend);
        }
    }
}
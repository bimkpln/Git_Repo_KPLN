using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace KPLN_Tools.ExecutableCommand
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandApartmentManagerShow : IExternalCommand
    {
        private static ApartmentManagerWindow _window;
        private static ApartmentManagerExternalController _controller;
        private static ApartmentPresetData _sessionPresetData;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;

                if (_window != null)
                {
                    ShowAndActivateExistingWindow();
                    return Result.Succeeded;
                }

                _controller = new ApartmentManagerExternalController();
                _window = new ApartmentManagerWindow(
                    DBWorkerService.CurrentDBUserSubDepartment.Id,
                    _controller,
                    _sessionPresetData != null ? _sessionPresetData.Clone() : null);

                _controller.AttachWindow(_window);
                _window.Closed += OnWindowClosed;

                new WindowInteropHelper(_window).Owner = uiapp.MainWindowHandle;

                _window.Show();
                _window.Activate();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private static void ShowAndActivateExistingWindow()
        {
            if (_window == null)
                return;

            if (!_window.IsVisible)
                _window.Show();

            if (_window.WindowState == WindowState.Minimized)
                _window.WindowState = WindowState.Normal;

            _window.Activate();
            _window.Topmost = true;
            _window.Topmost = false;
            _window.Focus();
        }

        private static void OnWindowClosed(object sender, EventArgs e)
        {
            if (_window != null)
                _sessionPresetData = _window.ApartmentPresetData != null
                    ? _window.ApartmentPresetData.Clone()
                    : null;

            if (_controller != null)
                _controller.DetachWindow();

            _controller = null;
            _window = null;
        }
    }

    internal class ApartmentManagerExternalController : IApartmentManagerExternalController
    {
        private readonly ApartmentManagerExternalHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public ApartmentManagerExternalController()
        {
            _handler = new ApartmentManagerExternalHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void AttachWindow(ApartmentManagerWindow window)
        {
            _handler.AttachWindow(window);
        }

        public void DetachWindow()
        {
            _handler.DetachWindow();
        }

        public void RequestPlaceApartment(int apartmentId)
        {
            _handler.PreparePlaceApartment(apartmentId);
            _externalEvent.Raise();
        }

        public void RequestConvertTo3D(ApartmentPresetData presetData)
        {
            _handler.PrepareConvertTo3D(presetData);
            _externalEvent.Raise();
        }

        public void RequestOpenApartmentPresets(ApartmentPresetData presetData)
        {
            _handler.PrepareOpenApartmentPresets(presetData);
            _externalEvent.Raise();
        }

        public void RequestUpdateApartmentMarks()
        {
            _handler.PrepareUpdateApartmentMarks();
            _externalEvent.Raise();
        }
    }

    internal class ApartmentManagerExternalHandler : IExternalEventHandler
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string ApartmentInstanceMarker = "[KPLN_APT_INSTANCE]";

        private static readonly Dictionary<string, string> _familyNameByPathCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private ApartmentManagerWindow _window;

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
        }

        private class ApartmentProcessState
        {
            public ElementId ApartmentId { get; set; }
            public bool HasPreparedWalls { get; set; }
            public bool HasCreatedWalls { get; set; }
            public bool HasCreatedRooms { get; set; }
            public bool HasInstalledDoors { get; set; }
            public int SkippedRoomsCount { get; set; }
            public int SkippedWallsCount { get; set; }
            public int SkippedDoorsCount { get; set; }
            public bool HasRoomAreaMismatch { get; set; }
            public bool HasDeletedRoomMismatch { get; set; }
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

        private class FamilyRoomMarker
        {
            public string RoomCategory { get; set; }
            public Transform LocalTransform { get; set; }
            public double WidthInternal { get; set; }
            public double DepthInternal { get; set; }
            public double ExpectedAreaInternal { get; set; }
        }

        private class FamilyDoorMarker
        {
            public string DoorTypeName2D { get; set; }
            public int DoorWidthMm { get; set; }
            public XYZ LocalPoint { get; set; }
            public string RoomCategory { get; set; }
            public string Comment { get; set; }
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

        private class PreparedRoomPlacement
        {
            public ElementId ApartmentId { get; set; }
            public string RoomName { get; set; }
            public XYZ InsertPoint { get; set; }
            public double ExpectedAreaInternal { get; set; }
        }

        private class PreparedApartmentRooms
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedRoomPlacement> Rooms { get; set; }

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

        private class DeletedRoomMismatchInfo
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
            OpenApartmentPresets,
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
                        ExecutePlaceApartment(uidoc, doc, _requestedApartmentId);
                        break;

                    case RequestType.ConvertTo3D:
                        string validationMessage;
                        if (!ValidatePresetBeforeConvertTo3D(doc, _requestedPresetData, out validationMessage))
                        {
                            TaskDialog.Show("Предупреждение", validationMessage);
                            return;
                        }

                        ExecuteConvertTo3D(doc, _requestedPresetData);
                        break;

                    case RequestType.OpenApartmentPresets:
                        ExecuteOpenApartmentPresets(uidoc, doc, _requestedPresetData);
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

        public void PrepareOpenApartmentPresets(ApartmentPresetData presetData)
        {
            _requestType = RequestType.OpenApartmentPresets;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        public void PrepareUpdateApartmentMarks()
        {
            _requestType = RequestType.UpdateApartmentMarks;
            _requestedApartmentId = 0;
            _requestedPresetData = null;
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





        private static long GetElementIdValue(ElementId id)
        {
#if Revit2024 || Debug2024
            return id.Value;
#else
            return id.IntegerValue;
#endif
        }

        private static double ConvertMmToInternal(int valueMm)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double ConvertInternalToMm(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        private static double ConvertInternalAreaToSquareMeters(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.SquareMeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_SQUARE_METERS);
#endif
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

            Parameter p =
                e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) ??
                e.LookupParameter("Комментарии") ??
                e.LookupParameter("Комментарий");

            if (p != null)
            {
                string value = p.AsString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                p =
                    typeElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS) ??
                    typeElem.LookupParameter("Комментарии") ??
                    typeElem.LookupParameter("Комментарий");

                if (p != null)
                {
                    string value = p.AsString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }
            }

            return null;
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
                    errors.Add("У квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " не найден параметр '" + parameterName + "'.");
                return false;
            }

            if (p.IsReadOnly)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " доступен только для чтения.");
                return false;
            }

            if (p.StorageType != StorageType.Double)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " имеет нечисловой тип.");
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
                TaskDialog.Show("KPLN. Менеджер квартир", "В проекте не найдено квартир, установленных данным плагином.");
                return;
            }

            int updatedCount = 0;
            int skippedCount = 0;
            List<string> errors = new List<string>();

            using (Transaction t = new Transaction(doc, "KPLN. Обновить марки квартир"))
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
                        errors.Add("Ошибка у квартиры ID = " + GetElementIdValue(apartmentFi.Id) + ": " + ex.Message);
                    }
                }

                t.Commit();
            }

            string message =
                "Обработано квартир: " + apartments.Count +
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

        private void ExecuteOpenApartmentPresets(UIDocument uidoc, Document doc, ApartmentPresetData currentPreset)
        {
            ViewPlan activeFloorPlan = doc.ActiveView as ViewPlan;
            if (activeFloorPlan == null || activeFloorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "Перед открытием преднастроек откройте план этажа.");
                return;
            }

            ApartmentPresetWindowContext context = BuildPresetWindowContext(doc, activeFloorPlan);

            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                ApartmentPresetsWindow wnd = new ApartmentPresetsWindow(
                    currentPreset != null ? currentPreset.Clone() : new ApartmentPresetData(),
                    context);

                wnd.Owner = _window;

                bool? res = wnd.ShowDialog();
                if (res == true && wnd.ResultPresetData != null)
                    _window.SetApartmentPresetData(wnd.ResultPresetData.Clone());
            }));
        }

        private ApartmentPresetWindowContext BuildPresetWindowContext(Document doc, ViewPlan activeFloorPlan)
        {
            ApartmentPresetWindowContext context = new ApartmentPresetWindowContext();

            List<ViewPlan> plans = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(x => !x.IsTemplate && x.ViewType == ViewType.FloorPlan && x.GenLevel != null)
                .OrderBy(x => x.Id != activeFloorPlan.Id)
                .ThenBy(x => x.Name)
                .ToList();

            foreach (ViewPlan plan in plans)
            {
                context.Plans.Add(new ApartmentPlanPresetOption
                {
                    PlanName = plan.Name,
                    IsResolved = false
                });
            }

            context.ResolvePlanData = delegate (string planName)
            {
                return BuildResolvedPlanPresetOption(doc, planName);
            };

            return context;
        }

        private ApartmentPlanPresetOption BuildResolvedPlanPresetOption(Document doc, string planName)
        {
            ApartmentPlanPresetOption option = new ApartmentPlanPresetOption();
            option.PlanName = planName;
            option.IsResolved = true;

            ViewPlan plan = FindTargetFloorPlan(doc, planName);
            if (plan == null)
            {
                option.LowerConstraintText = "";
                option.UpperConstraintText = "Неприсоединённая";
                option.RoomCategories = new List<string> { "Помещение" };
                return option;
            }

            option.LowerConstraintText = BuildLowerConstraintTextForPlan(doc, plan);
            option.UpperConstraintText = "Неприсоединённая";
            option.WallThicknesses = BuildWallThicknessesForPlan(doc, plan);
            option.WallTypeOptionsByThickness = BuildWallTypeOptionsByThicknessForPlan(doc, plan);
            option.RoomCategories = BuildRoomCategoriesForPlan(doc, plan);

            if (option.RoomCategories == null || option.RoomCategories.Count == 0)
                option.RoomCategories = new List<string> { "Помещение" };

            option.DoorRequirements = BuildDoorRequirementsForPlan(doc, plan);
            option.DoorTypeOptionsByRequirementKey = BuildDoorTypeOptionsByRequirementForPlan(doc, option.DoorRequirements);

            return option;
        }

        private string BuildLowerConstraintTextForPlan(Document doc, ViewPlan plan)
        {
            if (plan == null)
                return "Ещё не проставлено ни одного 2D-семейства";

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return "Ещё не проставлено ни одного 2D-семейства";

            HashSet<string> levelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance fi in apartments)
            {
                ElementId levelId = GetInstanceLevelId(fi);
                if (levelId == ElementId.InvalidElementId)
                    continue;

                Level lvl = doc.GetElement(levelId) as Level;
                if (lvl != null && !string.IsNullOrWhiteSpace(lvl.Name))
                    levelNames.Add(lvl.Name);
            }

            if (levelNames.Count == 0)
                return "Ещё не проставлено ни одного 2D-семейства";

            return string.Join(", ", levelNames.OrderBy(x => x));
        }

        private static ViewPlan FindTargetFloorPlan(Document doc, string selectedPlanName)
        {
            List<ViewPlan> plans = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(x => !x.IsTemplate && x.ViewType == ViewType.FloorPlan && x.GenLevel != null)
                .ToList();

            if (!string.IsNullOrWhiteSpace(selectedPlanName))
            {
                ViewPlan selected = plans.FirstOrDefault(x =>
                    string.Equals(x.Name, selectedPlanName, StringComparison.OrdinalIgnoreCase));

                if (selected != null)
                    return selected;
            }

            ViewPlan activePlan = doc.ActiveView as ViewPlan;
            if (activePlan != null && !activePlan.IsTemplate && activePlan.ViewType == ViewType.FloorPlan && activePlan.GenLevel != null)
                return activePlan;

            return null;
        }

        private List<string> BuildRoomCategoriesForPlan(Document doc, ViewPlan plan)
        {
            HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (plan == null || plan.GenLevel == null)
                return new List<string> { "Помещение" };

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesOnLevel(doc, plan.GenLevel);

            if (apartments.Count == 0)
                return new List<string> { "Помещение" };

            foreach (FamilyInstance apartment in apartments)
            {
                List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartment);

                foreach (FamilyInstance roomFi in rooms)
                {
                    string categoryName = GetRoomCategoryLabel(roomFi);
                    if (!string.IsNullOrWhiteSpace(categoryName))
                        result.Add(categoryName);
                }
            }

            if (result.Count == 0)
                return new List<string> { "Помещение" };

            return result.OrderBy(x => x).ToList();
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesForPlan(Document doc, ViewPlan plan)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            if (doc == null || plan == null)
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

                bool belongsToPlan = false;

                if (fi.OwnerViewId != ElementId.InvalidElementId && fi.OwnerViewId == plan.Id)
                    belongsToPlan = true;

                if (!belongsToPlan && plan.GenLevel != null)
                {
                    ElementId levelId = GetInstanceLevelId(fi);
                    if (levelId != ElementId.InvalidElementId && levelId == plan.GenLevel.Id)
                        belongsToPlan = true;
                }

                if (belongsToPlan)
                    result.Add(fi);
            }

            return result;
        }

        private static List<FamilyInstance> FindRoomSubComponents(Document doc, FamilyInstance apartmentInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (apartmentInstance == null)
                return result;

            ICollection<ElementId> subIds = apartmentInstance.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return result;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                string familyName = "";
                string typeName = "";

                if (subFi.Symbol != null)
                {
                    typeName = subFi.Symbol.Name ?? "";
                    if (subFi.Symbol.Family != null)
                        familyName = subFi.Symbol.Family.Name ?? "";
                }

                Category cat = subFi.Category;

                bool byFamily = familyName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;
                bool byType = typeName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;
                bool byCategory = cat != null && cat.Name.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0;

                if (byFamily || byType || byCategory)
                    result.Add(subFi);
            }

            return result;
        }

        private static string GetRoomCategoryLabel(FamilyInstance roomFi)
        {
            if (roomFi == null)
                return null;

            string[] candidateParams = new[]
            {
                "Категория помещения",
                "Категория",
                "Назначение",
                "Имя",
                "Наименование"
            };

            foreach (string paramName in candidateParams)
            {
                Parameter p = roomFi.LookupParameter(paramName);
                if (p != null)
                {
                    string value = p.AsString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }
            }

            if (roomFi.Symbol != null && !string.IsNullOrWhiteSpace(roomFi.Symbol.Name))
                return roomFi.Symbol.Name.Trim();

            if (roomFi.Symbol != null && roomFi.Symbol.Family != null && !string.IsNullOrWhiteSpace(roomFi.Symbol.Family.Name))
                return roomFi.Symbol.Family.Name.Trim();

            if (roomFi.Category != null && !string.IsNullOrWhiteSpace(roomFi.Category.Name))
                return roomFi.Category.Name.Trim();

            return "Помещение";
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesOnLevel(Document doc, Level level)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (FamilyInstance fi in instances)
            {
                Parameter pComment = fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pComment == null)
                    continue;

                string comment = pComment.AsString();
                if (string.IsNullOrWhiteSpace(comment))
                    continue;

                if (!comment.Contains(ApartmentInstanceMarker))
                    continue;

                ElementId levelId = GetInstanceLevelId(fi);
                if (levelId == ElementId.InvalidElementId)
                    continue;

                if (levelId == level.Id)
                    result.Add(fi);
            }

            return result;
        }

        private List<int> BuildWallThicknessesForPlan(Document doc, ViewPlan plan)
        {
            HashSet<int> result = new HashSet<int>();

            if (plan == null)
                return new List<int>();

            List<FamilyInstance> apartmentsOnPlan = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartmentsOnPlan.Count == 0)
                return new List<int>();

            foreach (FamilyInstance fi in apartmentsOnPlan)
            {
                double thicknessInternal;
                if (!TryGetApartmentWallThickness(fi, out thicknessInternal))
                    continue;

                int thicknessMm = (int)Math.Round(ConvertInternalToMm(thicknessInternal));
                if (thicknessMm > 0)
                    result.Add(thicknessMm);
            }

            return result.OrderBy(x => x).ToList();
        }

        private static bool TryGetApartmentWallThickness(FamilyInstance apartmentFi, out double thicknessInternal)
        {
            thicknessInternal = 0;

            if (apartmentFi == null)
                return false;

            Parameter p = apartmentFi.LookupParameter("Стены_Толщина");
            if (p == null)
                p = apartmentFi.LookupParameter("Стены толщина");

            if (p != null && p.StorageType == StorageType.Double)
            {
                thicknessInternal = p.AsDouble();
                return true;
            }

            Element typeElem = apartmentFi.Document.GetElement(apartmentFi.GetTypeId());
            if (typeElem != null)
            {
                p = typeElem.LookupParameter("Стены_Толщина");
                if (p == null)
                    p = typeElem.LookupParameter("Стены толщина");

                if (p != null && p.StorageType == StorageType.Double)
                {
                    thicknessInternal = p.AsDouble();
                    return true;
                }
            }

            return false;
        }

        private static double GetApartmentWallThickness(FamilyInstance apartmentFi)
        {
            if (apartmentFi == null)
                throw new ArgumentNullException("apartmentFi");

            Parameter p = apartmentFi.LookupParameter("Стены_Толщина");
            if (p == null)
                p = apartmentFi.LookupParameter("Стены толщина");

            if (p != null && p.StorageType == StorageType.Double)
                return p.AsDouble();

            Element typeElem = apartmentFi.Document.GetElement(apartmentFi.GetTypeId());
            if (typeElem != null)
            {
                p = typeElem.LookupParameter("Стены_Толщина");
                if (p == null)
                    p = typeElem.LookupParameter("Стены толщина");

                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            throw new Exception("Не найден параметр 'Стены_Толщина' у экземпляра или типа семейства квартиры.");
        }

        private Dictionary<int, List<string>> BuildWallTypeOptionsByThicknessForPlan(Document doc, ViewPlan plan)
        {
            Dictionary<int, List<string>> result = new Dictionary<int, List<string>>();

            List<int> thicknesses = BuildWallThicknessesForPlan(doc, plan);
            if (thicknesses.Count == 0)
                return result;

            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            foreach (int thicknessMm in thicknesses)
            {
                List<string> matchedNames = new List<string>();

                foreach (WallType wt in allWallTypes)
                {
                    if (wt == null || string.IsNullOrWhiteSpace(wt.Name))
                        continue;

                    int wallTypeThicknessMm;
                    if (!TryGetWallTypeThicknessMm(wt, out wallTypeThicknessMm))
                        continue;

                    if (wallTypeThicknessMm == thicknessMm)
                        matchedNames.Add(wt.Name);
                }

                result[thicknessMm] = matchedNames
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();
            }

            return result;
        }

        private static bool TryGetWallTypeThicknessMm(WallType wallType, out int thicknessMm)
        {
            thicknessMm = 0;

            if (wallType == null)
                return false;

            Parameter p = wallType.LookupParameter("Толщина");
            if (p != null && p.StorageType == StorageType.Double)
            {
                thicknessMm = (int)Math.Round(ConvertInternalToMm(p.AsDouble()));
                return thicknessMm > 0;
            }

            try
            {
                double width = wallType.Width;
                if (width > 0)
                {
                    thicknessMm = (int)Math.Round(ConvertInternalToMm(width));
                    return thicknessMm > 0;
                }
            }
            catch
            {
            }

            return false;
        }

        private List<ApartmentDoorRequirementOption> BuildDoorRequirementsForPlan(Document doc, ViewPlan plan)
        {
            Dictionary<string, ApartmentDoorRequirementOption> result =
                new Dictionary<string, ApartmentDoorRequirementOption>(StringComparer.OrdinalIgnoreCase);

            if (doc == null || plan == null)
                return new List<ApartmentDoorRequirementOption>();

            List<FamilyInstance> apartments = GetPlacedApartmentInstancesForPlan(doc, plan);
            if (apartments.Count == 0)
                return new List<ApartmentDoorRequirementOption>();

            HashSet<string> processedApartmentFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance apartmentFi in apartments)
            {
                if (apartmentFi == null || apartmentFi.Symbol == null || apartmentFi.Symbol.Family == null)
                    continue;

                Family apartmentFamily = apartmentFi.Symbol.Family;
                string apartmentFamilyKey = apartmentFamily.Name ?? "";
                if (string.IsNullOrWhiteSpace(apartmentFamilyKey))
                    apartmentFamilyKey = GetElementIdValue(apartmentFamily.Id).ToString();

                if (processedApartmentFamilies.Contains(apartmentFamilyKey))
                    continue;

                processedApartmentFamilies.Add(apartmentFamilyKey);

                List<FamilyRoomMarker> rooms = new List<FamilyRoomMarker>();
                List<FamilyDoorMarker> doors = new List<FamilyDoorMarker>();
                HashSet<string> visitedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                CollectApartmentFamilyMarkersRecursive(
                    doc,
                    apartmentFamily,
                    Transform.Identity,
                    rooms,
                    doors,
                    visitedFamilies);

                foreach (FamilyDoorMarker door in doors)
                {
                    string roomCategory = !string.IsNullOrWhiteSpace(door.RoomCategory)
                        ? door.RoomCategory.Trim()
                        : "-";

                    string key = ApartmentDoorRequirementOption.BuildKey(roomCategory, door.DoorTypeName2D, door.DoorWidthMm);

                    if (!result.ContainsKey(key))
                    {
                        result.Add(key, new ApartmentDoorRequirementOption
                        {
                            RoomCategory = roomCategory,
                            DoorTypeName2D = door.DoorTypeName2D,
                            WidthMm = door.DoorWidthMm
                        });
                    }
                }
            }

            return result.Values
                .OrderBy(x => x.RoomCategory)
                .ThenBy(x => x.WidthMm)
                .ThenBy(x => x.DoorTypeName2D)
                .ToList();
        }

        private static void CollectApartmentFamilyMarkersRecursive(Document ownerDoc, Family family, Transform accumulatedLocalTransform, List<FamilyRoomMarker> rooms,
            List<FamilyDoorMarker> doors, HashSet<string> visitedFamilies)
        {
            if (ownerDoc == null || family == null || accumulatedLocalTransform == null ||
                rooms == null || doors == null || visitedFamilies == null)
                return;

            string familyKey = family.Name ?? "";
            if (string.IsNullOrWhiteSpace(familyKey))
                familyKey = GetElementIdValue(family.Id).ToString();

            if (visitedFamilies.Contains(familyKey))
                return;

            visitedFamilies.Add(familyKey);

            Document familyDoc = null;

            try
            {
                familyDoc = ownerDoc.EditFamily(family);

                List<FamilyInstance> nestedInstances = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (FamilyInstance fi in nestedInstances)
                {
                    string familyName = "";
                    string typeName = "";
                    string categoryName = "";

                    if (fi.Symbol != null)
                    {
                        typeName = fi.Symbol.Name ?? "";
                        if (fi.Symbol.Family != null)
                            familyName = fi.Symbol.Family.Name ?? "";
                    }

                    if (fi.Category != null)
                        categoryName = fi.Category.Name ?? "";

                    Transform localTransform = fi.GetTransform();
                    if (localTransform == null)
                        localTransform = Transform.Identity;

                    Transform currentLocalTransform = accumulatedLocalTransform.Multiply(localTransform);

                    bool isRoomLike =
                        familyName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        typeName.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        (fi.Category != null && fi.Category.Name.IndexOf("Помещение", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (isRoomLike)
                    {
                        try
                        {
                            double width = GetRequiredLengthParam(fi, "Ширина", "Width");
                            double depth = GetRequiredLengthParam(fi, "Глубина", "Depth");

                            double expectedAreaInternal = 0;
                            TryGetAreaParamFromElementOrType(fi, out expectedAreaInternal, "КП_Р_Площадь", "КП_Р_ПЛОЩАДЬ");

                            rooms.Add(new FamilyRoomMarker
                            {
                                RoomCategory = GetRoomCategoryLabel(fi),
                                LocalTransform = currentLocalTransform,
                                WidthInternal = width,
                                DepthInternal = depth,
                                ExpectedAreaInternal = expectedAreaInternal
                            });
                        }
                        catch
                        {
                        }
                    }

                    bool isGenericModel =
                        string.Equals(categoryName, "Обобщенные модели", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(categoryName, "Обобщённые модели", StringComparison.OrdinalIgnoreCase);

                    bool isDoorFamily =
                        string.Equals(familyName, "Дверь", StringComparison.OrdinalIgnoreCase);

                    if (isGenericModel && isDoorFamily)
                    {
                        int widthMm;
                        if (TryGetDoorWidthMmFrom2DTypeName(typeName, out widthMm) && widthMm > 0)
                        {
                            string commentValue = GetCommentsValue(fi);

                            doors.Add(new FamilyDoorMarker
                            {
                                DoorTypeName2D = typeName,
                                DoorWidthMm = widthMm,
                                LocalPoint = currentLocalTransform.Origin,
                                Comment = commentValue,
                                RoomCategory = !string.IsNullOrWhiteSpace(commentValue) ? commentValue : null
                            });
                        }
                    }

                    Family nestedFamily = fi.Symbol != null ? fi.Symbol.Family : null;
                    if (nestedFamily != null)
                    {
                        CollectApartmentFamilyMarkersRecursive(familyDoc, nestedFamily, currentLocalTransform, rooms, doors, visitedFamilies);
                    }
                }
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
        }

        private static bool TryGetDoorWidthMmFrom2DTypeName(string typeName, out int widthMm)
        {
            widthMm = 0;

            if (string.IsNullOrWhiteSpace(typeName))
                return false;

            string normalized = typeName.Trim();

            int parsed;
            if (int.TryParse(normalized, out parsed))
            {
                widthMm = parsed;
                return widthMm > 0;
            }

            string firstToken = normalized
                .Split(new[] { ' ', 'x', 'X', 'х', 'Х', '-', '_', '/', '\\', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(firstToken) && int.TryParse(firstToken, out parsed))
            {
                widthMm = parsed;
                return widthMm > 0;
            }

            return false;
        }

        private Dictionary<string, List<string>> BuildDoorTypeOptionsByRequirementForPlan(Document doc, List<ApartmentDoorRequirementOption> requirements)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            if (requirements == null || requirements.Count == 0)
                return result;

            List<FamilySymbol> doorTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            foreach (ApartmentDoorRequirementOption requirement in requirements)
            {
                if (requirement == null || string.IsNullOrWhiteSpace(requirement.DoorTypeName2D) || requirement.WidthMm <= 0)
                    continue;

                List<string> matched = new List<string>();

                foreach (FamilySymbol symbol in doorTypes)
                {
                    int widthMm;
                    if (!TryGetProjectDoorTypeWidthMm(symbol, out widthMm))
                        continue;

                    if (widthMm != requirement.WidthMm)
                        continue;

                    string displayName = BuildDoorTypeDisplayName(symbol);
                    if (!string.IsNullOrWhiteSpace(displayName))
                        matched.Add(displayName);
                }

                result[requirement.Key] = matched
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();
            }

            return result;
        }

        private static bool TryGetProjectDoorTypeWidthMm(FamilySymbol symbol, out int widthMm)
        {
            widthMm = 0;
            if (symbol == null)
                return false;

            double widthInternal;
            if (TryGetLengthParamFromElementOrType(symbol, out widthInternal, "Ширина", "Width"))
            {
                widthMm = (int)Math.Round(ConvertInternalToMm(widthInternal));
                return widthMm > 0;
            }

            Parameter builtInParam = symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH);
            if (builtInParam != null && builtInParam.StorageType == StorageType.Double)
            {
                widthMm = (int)Math.Round(ConvertInternalToMm(builtInParam.AsDouble()));
                return widthMm > 0;
            }

            return false;
        }

        private static bool TryGetLengthParamFromElementOrType(Element e, out double valueInternal, params string[] paramNames)
        {
            valueInternal = 0;
            if (e == null)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    valueInternal = p.AsDouble();
                    return true;
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        valueInternal = p.AsDouble();
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetAreaParamFromElementOrType(Element e, out double valueInternal, params string[] paramNames)
        {
            valueInternal = 0;
            if (e == null)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    valueInternal = p.AsDouble();
                    return true;
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        valueInternal = p.AsDouble();
                        return true;
                    }
                }
            }

            return false;
        }

        private static string BuildDoorTypeDisplayName(FamilySymbol symbol)
        {
            if (symbol == null)
                return null;

            string familyName = symbol.Family != null ? symbol.Family.Name ?? "" : "";
            string typeName = symbol.Name ?? "";

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
                return familyName + " - " + typeName;

            if (!string.IsNullOrWhiteSpace(typeName))
                return typeName;

            return familyName;
        }

        private static FamilySymbol FindDoorSymbolByDisplayNameAndWidth(Document doc, string displayName, int widthMm)
        {
            if (doc == null || string.IsNullOrWhiteSpace(displayName))
                return null;

            List<FamilySymbol> doorTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            foreach (FamilySymbol symbol in doorTypes)
            {
                string currentName = BuildDoorTypeDisplayName(symbol);
                if (!string.Equals(currentName, displayName, StringComparison.OrdinalIgnoreCase))
                    continue;

                int symbolWidthMm;
                if (!TryGetProjectDoorTypeWidthMm(symbol, out symbolWidthMm))
                    continue;

                if (symbolWidthMm == widthMm)
                    return symbol;
            }

            return null;
        }





        private void ExecutePlaceApartment(UIDocument uidoc, Document doc, int id)
        {
            ViewPlan floorPlan = doc.ActiveView as ViewPlan;
            if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("Предупреждение", "Откройте план этажа перед размещением семейства.");
                return;
            }

            string familyPath = GetFamilyPathById(id);
            if (string.IsNullOrWhiteSpace(familyPath))
            {
                TaskDialog.Show("Ошибка", "Для ID = " + id + " не найден FPATH в базе.");
                return;
            }

            if (!File.Exists(familyPath))
            {
                TaskDialog.Show("Ошибка", "Файл семейства не найден:\n" + familyPath);
                return;
            }

            FamilySymbol symbol = null;

            using (Transaction t = new Transaction(doc, "Подготовить семейство квартиры"))
            {
                t.Start();

                Family family = LoadOrFindFamily(doc, familyPath);
                if (family == null)
                    throw new Exception("Не удалось загрузить или найти семейство в проекте.");

                symbol = GetFirstFamilySymbol(doc, family);
                if (symbol == null)
                    throw new Exception("У семейства не найдено ни одного типоразмера.");

                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }

                t.Commit();
            }

            int placedCount = 0;

            using (TransactionGroup tg = new TransactionGroup(doc, "Разместить семейство квартиры"))
            {
                tg.Start();

                while (true)
                {
                    XYZ insertPoint;

                    try
                    {
                        insertPoint = uidoc.Selection.PickPoint("Укажите точку вставки семейства. ESC - завершить.");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        break;
                    }

                    using (Transaction t = new Transaction(doc, "Разместить экземпляр квартиры"))
                    {
                        t.Start();

                        FamilyInstance placedInstance = PlaceFamilyInstance(doc, floorPlan, symbol, insertPoint);
                        if (placedInstance == null)
                            throw new Exception("Не удалось разместить экземпляр семейства.");

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
        }





        private static Family LoadOrFindFamily(Document doc, string familyPath)
        {
            if (doc == null)
                throw new ArgumentNullException("doc");

            if (string.IsNullOrWhiteSpace(familyPath))
                throw new ArgumentException("Не указан путь к семейству.", "familyPath");

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

        private static void RemoveApartmentInstanceMarker(Element e)
        {
            if (e == null)
                return;

            Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p == null || p.IsReadOnly)
                return;

            string oldValue = p.AsString();
            if (string.IsNullOrWhiteSpace(oldValue) || !oldValue.Contains(ApartmentInstanceMarker))
                return;

            string newValue = oldValue.Replace(ApartmentInstanceMarker, " ");
            while (newValue.Contains("  "))
                newValue = newValue.Replace("  ", " ");

            newValue = newValue.Trim();
            p.Set(newValue);
        }




        private bool ValidatePresetBeforeConvertTo3D(Document doc, ApartmentPresetData preset, out string validationMessage)
        {
            validationMessage = "";

            ApartmentPresetData effectivePreset = preset ?? new ApartmentPresetData
            {
                SelectedPlanName = "",
                LowerConstraint = "",
                UpperConstraint = "Неприсоединённая",
                BaseOffset = 0,
                WallHeight = 3000,
                WallTypeByThickness = new Dictionary<int, string>(),
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = new Dictionary<string, string>(),
                FamilyPostProcessAction = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            List<string> notFilled = new List<string>();
            List<string> otherProblems = new List<string>();

            if (effectivePreset.WallHeight <= 0)
                notFilled.Add("Неприсоединённая высота стены");

            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
            if (targetPlan == null)
            {
                otherProblems.Add("Не удалось определить план для построения стен.");
            }
            else
            {
                List<int> requiredThicknesses = BuildWallThicknessesForPlan(doc, targetPlan);

                if (requiredThicknesses.Count == 0)
                {
                    otherProblems.Add("На выбранном плане не найдено ни одного размещённого 2D-семейства квартиры.");
                }
                else
                {
                    foreach (int thickness in requiredThicknesses.OrderBy(x => x))
                    {
                        string selectedWallType = GetSelectedWallTypeNameForThickness(effectivePreset, thickness);

                        if (string.IsNullOrWhiteSpace(selectedWallType) ||
                            string.Equals(selectedWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        {
                            notFilled.Add("Стена (" + thickness + ")");
                            continue;
                        }

                        WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallType, thickness);
                        if (matchedWallType == null)
                        {
                            otherProblems.Add(
                                "Для толщины " + thickness + " мм выбран тип стены '" + selectedWallType + "', но такой тип не найден в проекте.");
                        }
                    }
                }

                List<ApartmentDoorRequirementOption> doorRequirements = BuildDoorRequirementsForPlan(doc, targetPlan);

                if (doorRequirements.Count > 0)
                {
                    foreach (ApartmentDoorRequirementOption req in doorRequirements
                        .Where(x => x != null)
                        .OrderBy(x => x.RoomCategory)
                        .ThenBy(x => x.WidthMm)
                        .ThenBy(x => x.DoorTypeName2D))
                    {
                        string selectedDoorTypeName = null;

                        if (effectivePreset.DoorsByRoomCategory != null)
                            effectivePreset.DoorsByRoomCategory.TryGetValue(req.Key, out selectedDoorTypeName);

                        if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                            string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        {
                            notFilled.Add("Дверь [" + req.RoomCategory + "] (" + req.WidthMm + ")");
                            continue;
                        }

                        FamilySymbol matchedDoorSymbol = FindDoorSymbolByDisplayNameAndWidth(doc, selectedDoorTypeName, req.WidthMm);
                        if (matchedDoorSymbol == null)
                        {
                            otherProblems.Add(
                                "Для категории '" + req.RoomCategory +
                                "', ширины " + req.WidthMm +
                                " мм выбран тип двери '" + selectedDoorTypeName +
                                "', но такой тип не найден в проекте.");
                        }
                    }
                }
            }

            if (notFilled.Count > 0 || otherProblems.Count > 0)
            {
                List<string> parts = new List<string>();
                parts.Add("Невозможно выполнить построение 3D.");

                if (notFilled.Count > 0)
                {
                    parts.Add("");
                    parts.Add("Не заполнено:");
                    parts.Add("- " + string.Join("\n- ", notFilled.Distinct().ToList()));
                }

                if (otherProblems.Count > 0)
                {
                    parts.Add("");
                    parts.Add("Ошибки:");
                    parts.Add("- " + string.Join("\n- ", otherProblems.Distinct().ToList()));
                }

                validationMessage = string.Join("\n", parts);
                return false;
            }

            return true;
        }






        private void ExecuteConvertTo3D(Document doc, ApartmentPresetData preset)
        {
            ApartmentPresetData effectivePreset = preset ?? new ApartmentPresetData
            {
                SelectedPlanName = "",
                LowerConstraint = "",
                UpperConstraint = "Неприсоединённая",
                BaseOffset = 0,
                WallHeight = 3000,
                WallTypeByThickness = new Dictionary<int, string>(),
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = new Dictionary<string, string>(),
                FamilyPostProcessAction = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
            if (targetPlan == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось определить план для построения.");
                return;
            }

            if (targetPlan.GenLevel == null)
            {
                TaskDialog.Show("Ошибка", "У выбранного плана не определён уровень.");
                return;
            }

            Level baseLevel = ResolveBaseLevelForPreset(doc, effectivePreset, targetPlan);
            Level topLevel = ResolveTopLevelForPreset(doc, effectivePreset);

            if (baseLevel == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось определить зависимость снизу.");
                return;
            }

            List<FamilyInstance> apartmentInstances = GetPlacedApartmentInstancesForPlan(doc, targetPlan);
            if (apartmentInstances.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "На выбранном плане не найдено ранее размещённых экземпляров квартир.");
                return;
            }

            List<string> debugMessages = new List<string>();
            List<PreparedApartmentWalls> preparedApartments = new List<PreparedApartmentWalls>();
            List<PreparedApartmentDoors> preparedDoorsByApartment = new List<PreparedApartmentDoors>();
            List<PreparedApartmentRooms> preparedRoomsByApartment = new List<PreparedApartmentRooms>();
            Dictionary<long, ApartmentProcessState> apartmentStates = new Dictionary<long, ApartmentProcessState>();

            double connectTol = ConvertMmToInternal(150);
            double intersectionTol = ConvertMmToInternal(10);

            List<ExistingWallLineInfo> existingWalls = GetExistingWallLinesOnLevel(doc, targetPlan.GenLevel.Id);

            foreach (FamilyInstance apartmentFi in apartmentInstances)
            {
                if (apartmentFi == null)
                    continue;

                ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentFi.Id);

                try
                {
                    double apartmentWallThicknessInternal = GetApartmentWallThickness(apartmentFi);
                    int apartmentWallThicknessMm = (int)Math.Round(ConvertInternalToMm(apartmentWallThicknessInternal));

                    if (apartmentWallThicknessMm <= 0)
                    {
                        debugMessages.Add("У квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " параметр 'Стены_Толщина' имеет некорректное значение.");
                        continue;
                    }

                    string selectedWallTypeName = GetSelectedWallTypeNameForThickness(effectivePreset, apartmentWallThicknessMm);
                    WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallTypeName, apartmentWallThicknessMm);

                    if (matchedWallType == null)
                    {
                        debugMessages.Add("Для квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " не найден тип стены '" + selectedWallTypeName + "' с толщиной " + apartmentWallThicknessMm + " мм.");
                        continue;
                    }

                    List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
                    if (roomInstances.Count == 0)
                    {
                        debugMessages.Add("Не найдены вложенные экземпляры 'Помещение' у экземпляра ID = " + GetElementIdValue(apartmentFi.Id));
                        continue;
                    }

                    List<Line> apartmentAxisLines = new List<Line>();

                    foreach (FamilyInstance roomFi in roomInstances)
                    {
                        try
                        {
                            int skippedWallsForApartment = state.SkippedWallsCount;

                            List<Line> roomAxisLines = BuildPreparedWallAxisLinesForSingleRoom(
                                roomFi,
                                apartmentWallThicknessInternal,
                                existingWalls,
                                connectTol,
                                intersectionTol,
                                debugMessages,
                                ref skippedWallsForApartment);

                            state.SkippedWallsCount = skippedWallsForApartment;

                            if (roomAxisLines == null || roomAxisLines.Count == 0)
                            {
                                state.SkippedRoomsCount++;
                                continue;
                            }

                            apartmentAxisLines.AddRange(roomAxisLines);
                        }
                        catch (Exception exRoom)
                        {
                            state.SkippedRoomsCount++;
                            debugMessages.Add("Ошибка обработки вложенного помещения ID = " + GetElementIdValue(roomFi.Id) + ": " + exRoom.Message);
                        }
                    }

                    apartmentAxisLines = MergeCollinearLines(apartmentAxisLines);
                    apartmentAxisLines = RemoveSegmentsOverlappingExistingWalls(apartmentAxisLines, existingWalls);
                    apartmentAxisLines = MergeCollinearLines(apartmentAxisLines);

                    if (apartmentAxisLines.Count == 0)
                    {
                        if (state.SkippedRoomsCount == 0)
                            state.SkippedRoomsCount = roomInstances.Count;

                        debugMessages.Add("Для квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " после покомнатной обработки не осталось осей стен.");
                    }
                    else
                    {
                        preparedApartments.Add(new PreparedApartmentWalls
                        {
                            ApartmentId = apartmentFi.Id,
                            WallType = matchedWallType,
                            ThicknessMm = apartmentWallThicknessMm,
                            AxisLines = apartmentAxisLines
                        });

                        state.HasPreparedWalls = true;
                    }

                    PreparedApartmentDoors preparedDoors = PrepareDoorsForApartment(doc, apartmentFi, effectivePreset, debugMessages);
                    preparedDoorsByApartment.Add(preparedDoors);

                    PreparedApartmentRooms preparedRooms = PrepareRoomsForApartment(doc, apartmentFi, debugMessages);
                    preparedRoomsByApartment.Add(preparedRooms);
                }
                catch (Exception exApartment)
                {
                    debugMessages.Add("Ошибка обработки квартиры ID = " + GetElementIdValue(apartmentFi.Id) + ": " + exApartment.Message);
                }
            }

            double baseOffsetInternal = ConvertMmToInternal(effectivePreset.BaseOffset);
            double wallHeightInternal = ConvertMmToInternal(effectivePreset.WallHeight > 0 ? effectivePreset.WallHeight : 3000);

            int totalDoorsPlanned = preparedDoorsByApartment
                .Where(x => x != null && x.Doors != null)
                .Sum(x => x.Doors.Count);

            int installedDoorsCount = 0;

            int totalRoomsPlanned = preparedRoomsByApartment
                .Where(x => x != null && x.Rooms != null)
                .Sum(x => x.Rooms.Count);

            int createdRoomsCount = 0;
            List<RoomAreaMismatchInfo> roomAreaMismatches = new List<RoomAreaMismatchInfo>();
            List<DeletedRoomMismatchInfo> deletedRoomMismatches = new List<DeletedRoomMismatchInfo>();

            if (preparedApartments.Count > 0)
            {
                using (Transaction t = new Transaction(doc, "KPLN. Построение стен по помещениям"))
                {
                    t.Start();

                    foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                    {
                        if (apartmentWalls == null || apartmentWalls.WallType == null || apartmentWalls.AxisLines == null)
                            continue;

                        ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);
                        List<Wall> createdWallsForApartment = new List<Wall>();

                        foreach (Line axis in apartmentWalls.AxisLines)
                        {
                            if (axis == null || axis.Length < 1e-6)
                                continue;

                            Wall wall = Wall.Create(doc, axis, apartmentWalls.WallType.Id, baseLevel.Id, wallHeightInternal, 0, false, false);
                            ApplyWallPresetParameters(wall, baseLevel, topLevel, baseOffsetInternal, wallHeightInternal);
                            createdWallsForApartment.Add(wall);
                        }




                        if (createdWallsForApartment.Count > 0)
                            state.HasCreatedWalls = true;

                        doc.Regenerate();

                        PreparedApartmentDoors apartmentDoors = preparedDoorsByApartment
                            .FirstOrDefault(x => x != null && x.ApartmentId == apartmentWalls.ApartmentId);

                        if (apartmentDoors != null && apartmentDoors.Doors != null && apartmentDoors.Doors.Count > 0)
                        {
                            CorrectWallDirectionsForApartmentBy2DDoors(
                                doc,
                                apartmentDoors,
                                createdWallsForApartment);

                            doc.Regenerate();

                            int installedDoorsForApartment = PlaceDoorsForApartment(
                                doc,
                                apartmentDoors,
                                createdWallsForApartment,
                                baseLevel);

                            installedDoorsCount += installedDoorsForApartment;

                            if (installedDoorsForApartment > 0)
                                state.HasInstalledDoors = true;

                            int skippedDoorsForApartment = apartmentDoors.Doors.Count - installedDoorsForApartment;
                            if (skippedDoorsForApartment > 0)
                                state.SkippedDoorsCount += skippedDoorsForApartment;
                        }

                        doc.Regenerate();






                        PreparedApartmentRooms apartmentRooms = preparedRoomsByApartment
                            .FirstOrDefault(x => x != null && x.ApartmentId == apartmentWalls.ApartmentId);

                        if (apartmentRooms != null && apartmentRooms.Rooms != null && apartmentRooms.Rooms.Count > 0)
                        {
                            int mismatchesBefore = roomAreaMismatches.Count;
                            int deletedBefore = deletedRoomMismatches.Count;

                            int createdRoomsForApartment = PlaceRoomsForApartment(doc, apartmentRooms, targetPlan.GenLevel, roomAreaMismatches, deletedRoomMismatches, createdWallsForApartment);
                            createdRoomsCount += createdRoomsForApartment;

                            if (createdRoomsForApartment > 0)
                                state.HasCreatedRooms = true;

                            int skippedRoomsForApartment = apartmentRooms.Rooms.Count - createdRoomsForApartment;
                            if (skippedRoomsForApartment > 0)
                                state.SkippedRoomsCount += skippedRoomsForApartment;

                            if (roomAreaMismatches.Count > mismatchesBefore)
                                state.HasRoomAreaMismatch = true;

                            if (deletedRoomMismatches.Count > deletedBefore)
                                state.HasDeletedRoomMismatch = true;
                        }
                    }

                    t.Commit();
                }
            }

            ApplyApartmentPostProcessAction(doc, apartmentInstances, effectivePreset.FamilyPostProcessAction, debugMessages);

            List<ApartmentExecutionReportItem> reportItems = BuildExecutionReportItems(apartmentStates, deletedRoomMismatches);

            int processedApartmentsCount = apartmentStates.Count;

            ShowExecutionReportWindow(targetPlan.Name, processedApartmentsCount, apartmentInstances.Count, createdRoomsCount, totalRoomsPlanned,
                installedDoorsCount, totalDoorsPlanned, reportItems);
        }





        private void ApplyApartmentPostProcessAction(Document doc, List<FamilyInstance> apartmentInstances, ApartmentFamilyPostProcessAction action, List<string> debugMessages)
        {
            if (doc == null || apartmentInstances == null || apartmentInstances.Count == 0)
                return;

            if (action == ApartmentFamilyPostProcessAction.Keep2DUnderlay)
                return;

            string transactionName =
                action == ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
                    ? "KPLN. Сохранение 2D-семейств с подложки"
                    : "KPLN. Полное удаление 2D-подложки";

            using (Transaction t = new Transaction(doc, transactionName))
            {
                t.Start();

                foreach (FamilyInstance apartmentFi in apartmentInstances)
                {
                    if (apartmentFi == null)
                        continue;

                    ElementId apartmentId = apartmentFi.Id;
                    if (apartmentId == ElementId.InvalidElementId)
                        continue;

                    if (action == ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay)
                    {
                        try
                        {
                            CopyFurnitureAndPlumbingFromApartmentUnderlay(doc, apartmentFi, debugMessages);
                        }
                        catch (Exception ex)
                        {
                            if (debugMessages != null)
                                debugMessages.Add(
                                    "Не удалось сохранить 2D-семейства из подложки квартиры ID = " +
                                    GetElementIdValue(apartmentId) + ": " + ex.Message);
                        }

                        TryDeleteSource2DApartmentInstance(doc, apartmentId, debugMessages);
                    }
                    else if (action == ApartmentFamilyPostProcessAction.Delete2DUnderlay)
                    {
                        TryDeleteSource2DApartmentInstance(doc, apartmentId, debugMessages);
                    }
                }

                t.Commit();
            }
        }





        private void ShowExecutionReportWindow(string planName, int processedApartments, int totalApartments, int createdRoomsCount,
            int totalRoomsPlanned, int installedDoorsCount, int totalDoorsPlanned, List<ApartmentExecutionReportItem> reportItems)
        {
            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                List<ApartmentExecutionReportItem> items = reportItems != null
                    ? reportItems.ToList()
                    : new List<ApartmentExecutionReportItem>();

                ApartmentExecutionReportItem summaryItem = new ApartmentExecutionReportItem
                {
                    ApartmentId = -1,
                    CustomHeaderText = "ОБЩИЙ РЕЗУЛЬТАТ"
                };

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "План: " + planName,
                    FontSize = 14,
                    FontWeight = FontWeights.SemiBold
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Экземпляров квартир обработано: " + processedApartments + " из " + totalApartments
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Помещений создано: " + createdRoomsCount + " из " + totalRoomsPlanned
                });

                summaryItem.Lines.Add(new ApartmentExecutionReportLine
                {
                    Text = "Дверей установлено: " + installedDoorsCount + " из " + totalDoorsPlanned
                });

                items.Insert(0, summaryItem);

                ApartmentExecutionReportWindow wnd = new ApartmentExecutionReportWindow(items);
                wnd.Owner = _window;
                wnd.ShowDialog();
            }));
        }

        private static List<ApartmentExecutionReportItem> BuildExecutionReportItems(Dictionary<long, ApartmentProcessState> apartmentStates,
            List<DeletedRoomMismatchInfo> deletedRoomMismatches)
        {
            List<ApartmentExecutionReportItem> result = new List<ApartmentExecutionReportItem>();

            if (apartmentStates == null || apartmentStates.Count == 0)
                return result;

            foreach (ApartmentProcessState state in apartmentStates.Values.OrderBy(x => GetElementIdValue(x.ApartmentId)))
            {
                if (state == null || state.ApartmentId == null || state.ApartmentId == ElementId.InvalidElementId)
                    continue;

                ApartmentExecutionReportItem reportItem = new ApartmentExecutionReportItem
                {
                    ApartmentId = GetElementIdValue(state.ApartmentId)
                };

                bool hasSkippedRooms = state.SkippedRoomsCount > 0;
                bool hasSkippedDoors = state.SkippedDoorsCount > 0;
                bool hasProblematicDeletedRooms = state.HasDeletedRoomMismatch || state.HasRoomAreaMismatch;
                bool onlyWallsSkipped = state.SkippedWallsCount > 0 && !hasSkippedRooms && !hasSkippedDoors && !hasProblematicDeletedRooms;
                bool fullySuccessful = state.SkippedWallsCount == 0 && !hasSkippedRooms && !hasSkippedDoors && !hasProblematicDeletedRooms;

                if (fullySuccessful || onlyWallsSkipped)
                {
                    reportItem.Lines.Add(new ApartmentExecutionReportLine
                    {
                        Text = "Квартира обработана успешно",
                        Foreground = System.Windows.Media.Brushes.Green,
                        FontSize = 14,
                        FontWeight = FontWeights.SemiBold
                    });
                }
                else
                {
                    if (state.SkippedWallsCount > 0 && hasSkippedRooms)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Часть стен не была построена",
                            Foreground = System.Windows.Media.Brushes.DarkOrange
                        });
                    }

                    if (hasSkippedRooms)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не построено помещений: " + state.SkippedRoomsCount,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });
                    }

                    if (hasSkippedDoors)
                    {
                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "Не построено дверей: " + state.SkippedDoorsCount,
                            Foreground = System.Windows.Media.Brushes.Red,
                            FontSize = 14,
                            FontWeight = FontWeights.SemiBold
                        });
                    }
                }

                if (deletedRoomMismatches != null)
                {
                    List<DeletedRoomMismatchInfo> deletedForApartment = deletedRoomMismatches
                        .Where(x => x != null && x.ApartmentId != null && GetElementIdValue(x.ApartmentId) == reportItem.ApartmentId)
                        .OrderBy(x => x.RoomName)
                        .ToList();

                    foreach (DeletedRoomMismatchInfo deletedItem in deletedForApartment)
                    {
                        string expectedText = ConvertInternalAreaToSquareMeters(deletedItem.ExpectedAreaInternal).ToString("0.##");
                        string actualText = ConvertInternalAreaToSquareMeters(deletedItem.ActualAreaInternal).ToString("0.##");

                        reportItem.Lines.Add(new ApartmentExecutionReportLine
                        {
                            Text = "-- Не совпавшее помещение: " +
                                   (string.IsNullOrWhiteSpace(deletedItem.RoomName) ? "Помещение" : deletedItem.RoomName) +
                                   " | 2D = " + expectedText + " м², 3D = " + actualText + " м²",
                            Foreground = System.Windows.Media.Brushes.DarkOrange
                        });
                    }
                }

                if (reportItem.Lines.Count == 0)
                {
                    reportItem.Lines.Add(new ApartmentExecutionReportLine
                    {
                        Text = "Квартира обработана успешно",
                        Foreground = System.Windows.Media.Brushes.Green,
                        FontSize = 14,
                        FontWeight = FontWeights.SemiBold
                    });
                }

                result.Add(reportItem);
            }

            return result;
        }

        private static ApartmentProcessState GetOrCreateApartmentState(Dictionary<long, ApartmentProcessState> states, ElementId apartmentId)
        {
            long key = GetElementIdValue(apartmentId);
            ApartmentProcessState state;

            if (!states.TryGetValue(key, out state))
            {
                state = new ApartmentProcessState
                {
                    ApartmentId = apartmentId
                };

                states.Add(key, state);
            }

            return state;
        }

        private List<Line> BuildPreparedWallAxisLinesForSingleRoom(FamilyInstance roomFi, double apartmentWallThicknessInternal, List<ExistingWallLineInfo> existingWalls,
            double connectTol, double intersectionTol, List<string> debugMessages, ref int skippedWallsForApartment)
        {
            if (roomFi == null)
                return new List<Line>();

            CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi);

            List<Line> wallAxisLines = BuildClosedWallAxisLinesFromRooms(
                new List<CurveLoop> { roomLoop },
                apartmentWallThicknessInternal,
                debugMessages);

            if (wallAxisLines.Count == 0)
                return new List<Line>();

            List<Line> preparedAxisLines = SnapNewLinesToExistingWalls(wallAxisLines, existingWalls, connectTol);
            preparedAxisLines = MergeCollinearLines(preparedAxisLines);
            preparedAxisLines = RemoveSegmentsOverlappingExistingWalls(preparedAxisLines, existingWalls);
            preparedAxisLines = MergeCollinearLines(preparedAxisLines);

            if (preparedAxisLines.Count == 0)
                return new List<Line>();

            List<Line> finalAxisLines = new List<Line>();

            foreach (Line axis in preparedAxisLines)
            {
                if (axis == null || axis.Length < 1e-6)
                    continue;

                if (IntersectsExistingWalls(axis, apartmentWallThicknessInternal, existingWalls, intersectionTol))
                {
                    skippedWallsForApartment++;
                    continue;
                }

                finalAxisLines.Add(axis);
            }

            return MergeCollinearLines(finalAxisLines);
        }

        private static bool TryDeleteSource2DApartmentInstance(Document doc, ElementId apartmentId, List<string> debugMessages)
        {
            if (doc == null || apartmentId == null || apartmentId == ElementId.InvalidElementId)
                return false;

            try
            {
                Element element = doc.GetElement(apartmentId);
                if (element == null)
                    return false;

                doc.Delete(apartmentId);
                return true;
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось удалить 2D-квартиру ID = " + GetElementIdValue(apartmentId) + ": " + ex.Message);

                return false;
            }
        }

        private static bool IsFurnitureOrPlumbingCategory(Category category)
        {
            if (category == null)
                return false;

            BuiltInCategory bic;
            try
            {
                bic = (BuiltInCategory)category.Id.IntegerValue;
            }
            catch
            {
                return false;
            }

            return bic == BuiltInCategory.OST_Furniture ||
                   bic == BuiltInCategory.OST_PlumbingFixtures;
        }

        private static List<FamilyInstance> FindFurnitureAndPlumbingSubComponentsRecursive(Document doc, FamilyInstance rootInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (doc == null || rootInstance == null)
                return result;

            CollectFurnitureAndPlumbingSubComponentsRecursive(doc, rootInstance, result);
            return result;
        }

        private static void CollectFurnitureAndPlumbingSubComponentsRecursive(
            Document doc,
            FamilyInstance current,
            List<FamilyInstance> result)
        {
            if (doc == null || current == null)
                return;

            ICollection<ElementId> subIds = current.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                if (IsFurnitureOrPlumbingCategory(subFi.Category))
                {
                    result.Add(subFi);
                    continue;
                }

                CollectFurnitureAndPlumbingSubComponentsRecursive(doc, subFi, result);
            }
        }

        private static Level ResolvePlacementLevelForNestedInstance(Document doc, FamilyInstance nestedFi, FamilyInstance apartmentFi)
        {
            if (doc == null)
                return null;

            ElementId nestedLevelId = GetInstanceLevelId(nestedFi);
            if (nestedLevelId != ElementId.InvalidElementId)
            {
                Level nestedLevel = doc.GetElement(nestedLevelId) as Level;
                if (nestedLevel != null)
                    return nestedLevel;
            }

            ElementId apartmentLevelId = GetInstanceLevelId(apartmentFi);
            if (apartmentLevelId != ElementId.InvalidElementId)
            {
                Level apartmentLevel = doc.GetElement(apartmentLevelId) as Level;
                if (apartmentLevel != null)
                    return apartmentLevel;
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .FirstOrDefault();
        }

        private static double GetRotationAngleOnXY(FamilyInstance fi)
        {
            if (fi == null)
                return 0.0;

            Transform tr = fi.GetTransform();
            if (tr == null)
                return 0.0;

            XYZ basisX = tr.BasisX;
            if (basisX == null)
                return 0.0;

            return Math.Atan2(basisX.Y, basisX.X);
        }

        private void CopyFurnitureAndPlumbingFromApartmentUnderlay(Document doc, FamilyInstance apartmentFi, List<string> debugMessages)
        {
            if (doc == null || apartmentFi == null)
                return;

            List<FamilyInstance> nestedItems = FindFurnitureAndPlumbingSubComponentsRecursive(doc, apartmentFi);
            if (nestedItems == null || nestedItems.Count == 0)
                return;

            foreach (FamilyInstance nestedFi in nestedItems)
            {
                if (nestedFi == null)
                    continue;

                try
                {
                    FamilySymbol symbol = nestedFi.Symbol;
                    if (symbol == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + " не найден тип.");
                        continue;
                    }

                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                        doc.Regenerate();
                    }

                    Transform tr = nestedFi.GetTransform();
                    if (tr == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + " не найден Transform.");
                        continue;
                    }

                    XYZ insertPoint = tr.Origin;
                    if (insertPoint == null)
                        continue;

                    Level level = ResolvePlacementLevelForNestedInstance(doc, nestedFi, apartmentFi);
                    if (level == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не найден уровень для вложенного элемента ID = " + GetElementIdValue(nestedFi.Id));
                        continue;
                    }

                    FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;
                    FamilyInstance created = null;

                    switch (placementType)
                    {
                        case FamilyPlacementType.ViewBased:
                            break;

                        case FamilyPlacementType.OneLevelBased:
                        case FamilyPlacementType.OneLevelBasedHosted:
                        case FamilyPlacementType.WorkPlaneBased:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;

                        default:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;
                    }

                    if (created == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не удалось создать экземпляр для вложенного элемента ID = " + GetElementIdValue(nestedFi.Id));
                        continue;
                    }

                    double angle = GetRotationAngleOnXY(nestedFi);
                    if (Math.Abs(angle) > 1e-9)
                    {
                        Line axis = Line.CreateBound(insertPoint, insertPoint + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, created.Id, axis, angle);
                    }
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Ошибка копирования вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + ": " + ex.Message);
                }
            }
        }













        private static void DeleteIsolatedCreatedWalls(Document doc, List<Wall> createdWallsForApartment)
        {
            if (doc == null || createdWallsForApartment == null || createdWallsForApartment.Count == 0)
                return;

            double tol = ConvertMmToInternal(10);
            List<Wall> validWalls = createdWallsForApartment
                .Where(x => x != null && x.IsValidObject)
                .ToList();

            if (validWalls.Count == 0)
                return;

            List<ElementId> wallsToDelete = new List<ElementId>();

            foreach (Wall currentWall in validWalls)
            {
                if (currentWall == null || !currentWall.IsValidObject)
                    continue;

                LocationCurve currentLc = currentWall.Location as LocationCurve;
                if (currentLc == null)
                    continue;

                Line currentLine = currentLc.Curve as Line;
                if (currentLine == null || currentLine.Length < 1e-9)
                    continue;

                XYZ a0 = currentLine.GetEndPoint(0);
                XYZ a1 = currentLine.GetEndPoint(1);

                bool hasConnection = false;

                foreach (Wall otherWall in validWalls)
                {
                    if (otherWall == null || !otherWall.IsValidObject || otherWall.Id == currentWall.Id)
                        continue;

                    LocationCurve otherLc = otherWall.Location as LocationCurve;
                    if (otherLc == null)
                        continue;

                    Line otherLine = otherLc.Curve as Line;
                    if (otherLine == null || otherLine.Length < 1e-9)
                        continue;

                    XYZ b0 = otherLine.GetEndPoint(0);
                    XYZ b1 = otherLine.GetEndPoint(1);

                    bool touchesByEndpoints =
                        Distance2D(a0, b0) <= tol ||
                        Distance2D(a0, b1) <= tol ||
                        Distance2D(a1, b0) <= tol ||
                        Distance2D(a1, b1) <= tol;

                    XYZ intersection;
                    bool intersects = TryIntersectSegments2D(a0, a1, b0, b1, out intersection, tol);

                    if (touchesByEndpoints || intersects)
                    {
                        hasConnection = true;
                        break;
                    }
                }

                if (!hasConnection)
                    wallsToDelete.Add(currentWall.Id);
            }

            if (wallsToDelete.Count > 0)
            {
                doc.Delete(wallsToDelete);

                createdWallsForApartment.RemoveAll(x =>
                    x == null ||
                    !x.IsValidObject ||
                    wallsToDelete.Any(id => id == x.Id));
            }
        }


        private static XYZ GetRoomCenterPoint(FamilyInstance roomFi)
        {
            if (roomFi == null)
                return null;

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                return null;

            return tr.Origin;
        }

        private static FamilyInstance FindBestMatchingRoomForDoor(
            FamilyInstance apartmentFi,
            string roomCategory,
            XYZ doorPoint,
            Document doc)
        {
            List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartmentFi);
            if (rooms == null || rooms.Count == 0)
                return null;

            IEnumerable<FamilyInstance> filteredRooms = rooms;

            if (!string.IsNullOrWhiteSpace(roomCategory) && roomCategory != "-")
            {
                List<FamilyInstance> exactRooms = rooms
                    .Where(x => string.Equals(
                        GetRoomCategoryLabel(x),
                        roomCategory,
                        StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (exactRooms.Count > 0)
                    filteredRooms = exactRooms;
            }

            FamilyInstance bestRoom = null;
            double bestDistance = double.MaxValue;

            foreach (FamilyInstance roomFi in filteredRooms)
            {
                XYZ center = GetRoomCenterPoint(roomFi);
                if (center == null)
                    continue;

                double dist = Distance2D(center, doorPoint);
                if (dist < bestDistance)
                {
                    bestDistance = dist;
                    bestRoom = roomFi;
                }
            }

            return bestRoom;
        }

        private static Line GetWallAxisLine(Wall wall)
        {
            if (wall == null)
                return null;

            LocationCurve lc = wall.Location as LocationCurve;
            if (lc == null)
                return null;

            return lc.Curve as Line;
        }

        private static XYZ GetClosestPointOnRoomRectangle(FamilyInstance roomFi, XYZ worldPoint)
        {
            if (roomFi == null || worldPoint == null)
                return null;

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                return null;

            Transform inv = tr.Inverse;
            if (inv == null)
                return null;

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            XYZ localPoint = inv.OfPoint(worldPoint);

            double clampedX = Math.Max(-halfW, Math.Min(halfW, localPoint.X));
            double clampedY = Math.Max(-halfD, Math.Min(halfD, localPoint.Y));

            XYZ localClosest = new XYZ(clampedX, clampedY, 0);
            return tr.OfPoint(localClosest);
        }

        private void CorrectWallDirectionsForApartmentBy2DDoors(
            Document doc,
            PreparedApartmentDoors apartmentDoors,
            List<Wall> createdWallsForApartment)
        {
            if (doc == null || apartmentDoors == null || apartmentDoors.Doors == null || createdWallsForApartment == null)
                return;

            double maxDistanceToWallAxis = ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = false;

                if (preparedDoor == null || preparedDoor.InsertPoint == null || preparedDoor.RelatedRoom2D == null)
                    continue;

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;

                bool foundHost = TryFindBestHostWallForDoor(
                    preparedDoor.InsertPoint,
                    createdWallsForApartment,
                    maxDistanceToWallAxis,
                    out hostWall,
                    out projectedPoint,
                    out distanceToWallAxis);

                if (!foundHost || hostWall == null || projectedPoint == null)
                    continue;

                Line wallAxis = GetWallAxisLine(hostWall);
                if (wallAxis == null)
                    continue;

                XYZ wallDir = Normalize2D(wallAxis.GetEndPoint(1) - wallAxis.GetEndPoint(0));
                if (wallDir == null)
                    continue;

                XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0);

                XYZ roomPoint = GetClosestPointOnRoomRectangle(preparedDoor.RelatedRoom2D, preparedDoor.InsertPoint);
                if (roomPoint == null)
                    continue;

                XYZ toRoom = roomPoint - projectedPoint;
                double sign = Dot2D(toRoom, wallNormal);


                if (sign > 0)
                {
                    preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = true;

                    LocationCurve lc = hostWall.Location as LocationCurve;
                    if (lc == null)
                        continue;

                    Line reversedAxis = Line.CreateBound(
                        wallAxis.GetEndPoint(1),
                        wallAxis.GetEndPoint(0));

                    lc.Curve = reversedAxis;
                }
            }
        }




        private int PlaceDoorsForApartment(Document doc, PreparedApartmentDoors apartmentDoors, List<Wall> createdWallsForApartment, Level baseLevel)
        {
            if (doc == null || apartmentDoors == null || createdWallsForApartment == null || baseLevel == null)
                return 0;

            if (apartmentDoors.Doors == null || apartmentDoors.Doors.Count == 0)
                return 0;

            int installedCount = 0;
            double maxDistanceToWallAxis = ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                if (preparedDoor == null || preparedDoor.DoorSymbol == null || preparedDoor.InsertPoint == null)
                    continue;

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;

                bool foundHost = TryFindBestHostWallForDoor(
                    preparedDoor.InsertPoint,
                    createdWallsForApartment,
                    maxDistanceToWallAxis,
                    out hostWall,
                    out projectedPoint,
                    out distanceToWallAxis);

                if (!foundHost || hostWall == null || projectedPoint == null)
                    continue;

                FamilySymbol symbolToPlace = preparedDoor.DoorSymbol;

                if (preparedDoor.RequiresOppositeDoorTypeAfterWallFlip)
                    symbolToPlace = GetOppositeDoorSymbol(doc, symbolToPlace);

                if (symbolToPlace == null)
                    continue;

                if (!symbolToPlace.IsActive)
                {
                    symbolToPlace.Activate();
                    doc.Regenerate();
                }

                FamilyInstance createdDoor = doc.Create.NewFamilyInstance(
                    projectedPoint,
                    symbolToPlace,
                    hostWall,
                    baseLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                if (createdDoor != null)
                    installedCount++;
            }

            return installedCount;
        }

        private bool TryFindBestHostWallForDoor(XYZ doorPoint, List<Wall> candidateWalls, double maxDistanceToWallAxis, out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            bestWall = null;
            bestProjectedPoint = null;
            bestDistance = double.MaxValue;

            if (doorPoint == null || candidateWalls == null || candidateWalls.Count == 0)
                return false;

            foreach (Wall wall in candidateWalls)
            {
                if (wall == null)
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line wallLine = lc.Curve as Line;
                if (wallLine == null || wallLine.Length < 1e-9)
                    continue;

                XYZ projectedPoint;
                double distance;

                if (!TryProjectPointToSegment2D(doorPoint, wallLine, out projectedPoint, out distance))
                    continue;

                if (distance > maxDistanceToWallAxis)
                    continue;

                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestWall = wall;
                    bestProjectedPoint = projectedPoint;
                }
            }

            return bestWall != null && bestProjectedPoint != null;
        }

        private bool TryProjectPointToSegment2D(XYZ point, Line segment, out XYZ projectedPoint, out double distance)
        {
            projectedPoint = null;
            distance = double.MaxValue;

            if (point == null || segment == null)
                return false;

            XYZ a = segment.GetEndPoint(0);
            XYZ b = segment.GetEndPoint(1);
            XYZ ab = b - a;

            double len2 = ab.X * ab.X + ab.Y * ab.Y;
            if (len2 < 1e-12)
                return false;

            double t = ((point.X - a.X) * ab.X + (point.Y - a.Y) * ab.Y) / len2;

            if (t < 0.0)
                t = 0.0;
            else if (t > 1.0)
                t = 1.0;

            double x = a.X + ab.X * t;
            double y = a.Y + ab.Y * t;
            double z = a.Z + (b.Z - a.Z) * t;

            projectedPoint = new XYZ(x, y, z);
            distance = Distance2D(point, projectedPoint);

            return true;
        }

        private static Level ResolveBaseLevelForPreset(Document doc, ApartmentPresetData preset, ViewPlan targetPlan)
        {
            if (preset != null && !string.IsNullOrWhiteSpace(preset.LowerConstraint))
            {
                string[] parts = preset.LowerConstraint
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToArray();

                foreach (string levelName in parts)
                {
                    Level lvl = FindLevelByName(doc, levelName);
                    if (lvl != null)
                        return lvl;
                }
            }

            if (targetPlan != null && targetPlan.GenLevel != null)
                return targetPlan.GenLevel;

            return null;
        }

        private static Level ResolveTopLevelForPreset(Document doc, ApartmentPresetData preset)
        {
            if (preset == null)
                return null;

            if (string.IsNullOrWhiteSpace(preset.UpperConstraint))
                return null;

            if (string.Equals(preset.UpperConstraint, "Неприсоединённая", StringComparison.OrdinalIgnoreCase))
                return null;

            return FindLevelByName(doc, preset.UpperConstraint.Trim());
        }

        private static Level FindLevelByName(Document doc, string levelName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(levelName))
                return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(x => string.Equals(x.Name, levelName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetSelectedWallTypeNameForThickness(ApartmentPresetData preset, int thicknessMm)
        {
            if (preset == null || preset.WallTypeByThickness == null || preset.WallTypeByThickness.Count == 0)
                return null;

            string value;
            if (preset.WallTypeByThickness.TryGetValue(thicknessMm, out value))
                return value;

            return null;
        }

        private static WallType FindWallTypeByExactSelectionAndThickness(Document doc, string selectedWallTypeName, int thicknessMm)
        {
            if (doc == null || string.IsNullOrWhiteSpace(selectedWallTypeName))
                return null;

            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            WallType exact = allWallTypes.FirstOrDefault(x => string.Equals(x.Name, selectedWallTypeName, StringComparison.OrdinalIgnoreCase));

            if (exact != null)
            {
                int exactThicknessMm;
                if (TryGetWallTypeThicknessMm(exact, out exactThicknessMm) && exactThicknessMm == thicknessMm)
                    return exact;
            }

            foreach (WallType wt in allWallTypes)
            {
                if (wt == null || !string.Equals(wt.Name, selectedWallTypeName, StringComparison.OrdinalIgnoreCase))
                    continue;

                int wallTypeThicknessMm;
                if (!TryGetWallTypeThicknessMm(wt, out wallTypeThicknessMm))
                    continue;

                if (wallTypeThicknessMm == thicknessMm)
                    return wt;
            }

            return null;
        }

        private FamilySymbol GetOppositeDoorSymbol(Document doc, FamilySymbol currentSymbol)
        {
            if (doc == null || currentSymbol == null || currentSymbol.Family == null)
                return currentSymbol;

            string currentTypeName = currentSymbol.Name ?? "";
            DoorOpeningMarker currentMarker = GetDoorOpeningMarkerFromTypeName(currentTypeName);

            if (currentMarker == DoorOpeningMarker.None)
                return currentSymbol;

            DoorOpeningMarker oppositeMarker =
                currentMarker == DoorOpeningMarker.Left
                    ? DoorOpeningMarker.Right
                    : DoorOpeningMarker.Left;

            string targetTypeName = ReplaceDoorOpeningMarker(currentTypeName, oppositeMarker);

            FamilySymbol result = FindFamilySymbolByTypeName(doc, currentSymbol.Family, targetTypeName);
            if (result != null)
                return result;

            if (oppositeMarker == DoorOpeningMarker.Right)
            {
                string altRightTypeName = ReplaceDoorOpeningMarker(currentTypeName, DoorOpeningMarker.RightAlt);
                result = FindFamilySymbolByTypeName(doc, currentSymbol.Family, altRightTypeName);
                if (result != null)
                    return result;
            }

            return currentSymbol;
        }

        private static List<ExistingWallLineInfo> GetExistingWallLinesOnLevel(Document doc, ElementId levelId)
        {
            List<ExistingWallLineInfo> result = new List<ExistingWallLineInfo>();

            IEnumerable<Wall> walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>();

            foreach (Wall wall in walls)
            {
                if (wall == null)
                    continue;

                if (wall.LevelId != levelId)
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line line = lc.Curve as Line;
                if (line == null || line.Length < 1e-9)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                double thicknessInternal = 0;
                try
                {
                    thicknessInternal = wall.Width;
                }
                catch
                {
                    thicknessInternal = 0;
                }

                result.Add(new ExistingWallLineInfo
                {
                    WallId = wall.Id,
                    P0 = p0,
                    P1 = p1,
                    Dir = dir,
                    Z = 0.5 * (p0.Z + p1.Z),
                    ThicknessInternal = thicknessInternal
                });
            }

            return result;
        }

        private static void ApplyWallPresetParameters(Wall wall, Level baseLevel, Level topLevel, double baseOffsetInternal, double unconnectedHeightInternal)
        {
            if (wall == null)
                return;

            Parameter pBaseConstraint = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
            if (pBaseConstraint != null && !pBaseConstraint.IsReadOnly && baseLevel != null)
                pBaseConstraint.Set(baseLevel.Id);

            Parameter pBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
            if (pBaseOffset != null && !pBaseOffset.IsReadOnly)
                pBaseOffset.Set(baseOffsetInternal);

            Parameter pTopConstraint = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            Parameter pUnconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);

            if (topLevel != null)
            {
                if (pTopConstraint != null && !pTopConstraint.IsReadOnly)
                    pTopConstraint.Set(topLevel.Id);
            }
            else
            {
                if (pUnconnectedHeight != null && !pUnconnectedHeight.IsReadOnly)
                    pUnconnectedHeight.Set(unconnectedHeightInternal);
            }
        }

        private PreparedApartmentDoors PrepareDoorsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, List<string> debugMessages)
        {
            PreparedApartmentDoors result = new PreparedApartmentDoors();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyInstance> doorInstances = FindDoorSubComponentsRecursive(doc, apartmentFi);

            foreach (FamilyInstance doorFi in doorInstances)
            {
                if (doorFi == null)
                    continue;

                string typeName = doorFi.Symbol != null ? doorFi.Symbol.Name ?? "" : "";

                int widthMm;
                if (!TryGetDoorWidthMmFrom2DTypeName(typeName, out widthMm) || widthMm <= 0)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не удалось определить ширину 2D-двери у экземпляра ID = " + GetElementIdValue(doorFi.Id));
                    continue;
                }

                string roomCategory = GetCommentsValue(doorFi);
                if (string.IsNullOrWhiteSpace(roomCategory))
                    roomCategory = "-";

                string presetKey = ApartmentDoorRequirementOption.BuildKey(roomCategory, typeName, widthMm);

                string selectedDoorTypeName = null;
                if (preset != null && preset.DoorsByRoomCategory != null)
                    preset.DoorsByRoomCategory.TryGetValue(presetKey, out selectedDoorTypeName);

                if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                    string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                {
                    if (debugMessages != null)
                        debugMessages.Add("Для двери [" + roomCategory + "] (" + widthMm + ") не выбран тип двери проекта.");
                    continue;
                }

                FamilySymbol baseDoorSymbol = FindDoorSymbolByDisplayNameAndWidth(doc, selectedDoorTypeName, widthMm);
                if (baseDoorSymbol == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не найден тип двери проекта '" + selectedDoorTypeName + "' с шириной " + widthMm + " мм.");
                    continue;
                }

                FamilySymbol resolvedDoorSymbol = ResolveDoorSymbolForPlacement(doc, doorFi, baseDoorSymbol, debugMessages);
                if (resolvedDoorSymbol == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add(
                            "Не удалось определить итоговый тип 3D-двери для 2D-двери ID = " +
                            GetElementIdValue(doorFi.Id) + ".");
                    continue;
                }

                Transform doorTransform = doorFi.GetTransform();
                if (doorTransform == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не удалось получить Transform для 2D-двери ID = " + GetElementIdValue(doorFi.Id));
                    continue;
                }

                XYZ insertPointInProject = doorTransform.Origin;

                FamilyInstance matchedRoom = FindBestMatchingRoomForDoor(
                    apartmentFi,
                    roomCategory,
                    insertPointInProject,
                    doc);

                XYZ expectedRoomPoint = matchedRoom != null
                    ? GetRoomCenterPoint(matchedRoom)
                    : null;

                result.Doors.Add(new PreparedDoorPlacement
                {
                    ApartmentId = apartmentFi.Id,
                    Door2DId = doorFi.Id,
                    RoomCategory = roomCategory,
                    DoorWidthMm = widthMm,
                    SelectedDoorTypeName = BuildDoorTypeDisplayName(resolvedDoorSymbol),
                    DoorSymbol = resolvedDoorSymbol,
                    InsertPoint = insertPointInProject,
                    RelatedRoom2D = matchedRoom,
                    RequiresOppositeDoorTypeAfterWallFlip = false
                });
            }

            return result;
        }

        private static DoorOpeningMarker GetDoorOpeningMarkerFromTypeName(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return DoorOpeningMarker.None;

            string name = typeName.Trim();

            if (name.IndexOf("(Л ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.Left;

            if (name.IndexOf("(Пр ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.RightAlt;

            if (name.IndexOf("(П ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.Right;

            return DoorOpeningMarker.None;
        }

        private static string ReplaceDoorOpeningMarker(string typeName, DoorOpeningMarker newMarker)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return typeName;

            string replacement = "";

            switch (newMarker)
            {
                case DoorOpeningMarker.Left:
                    replacement = "(Л ";
                    break;
                case DoorOpeningMarker.Right:
                    replacement = "(П ";
                    break;
                case DoorOpeningMarker.RightAlt:
                    replacement = "(Пр ";
                    break;
                default:
                    return typeName;
            }

            string result = typeName;

            if (result.IndexOf("(Л ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(Л ", replacement);

            if (result.IndexOf("(Пр ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(Пр ", replacement);

            if (result.IndexOf("(П ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(П ", replacement);

            return typeName;
        }

        private static string ReplaceOrdinalIgnoreCase(string source, string oldValue, string newValue)
        {
            int index = source.IndexOf(oldValue, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
                return source;

            return source.Substring(0, index) + newValue + source.Substring(index + oldValue.Length);
        }

        private static Parameter FindDoorLeftOpeningParameter(Element e)
        {
            if (e == null)
                return null;

            Parameter p = e.LookupParameter("КП_О_Левое открывание");
            if (p != null)
                return p;

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                p = typeElem.LookupParameter("КП_О_Левое открывание");
                if (p != null)
                    return p;
            }

            return null;
        }

        private static bool SetYesNoParameter(Parameter p, bool value)
        {
            if (p == null || p.IsReadOnly)
                return false;

            if (p.StorageType != StorageType.Integer)
                return false;

            p.Set(value ? 1 : 0);
            return true;
        }

        private static List<FamilySymbol> GetAllSymbolsOfFamily(Document doc, Family family)
        {
            List<FamilySymbol> result = new List<FamilySymbol>();
            if (doc == null || family == null)
                return result;

            ISet<ElementId> ids = family.GetFamilySymbolIds();
            if (ids == null || ids.Count == 0)
                return result;

            foreach (ElementId id in ids)
            {
                FamilySymbol symbol = doc.GetElement(id) as FamilySymbol;
                if (symbol != null)
                    result.Add(symbol);
            }

            return result;
        }






        private DoorTypeMirrorEnsureResult EnsureDoorMirrorTypeExists(Document doc, FamilySymbol sourceSymbol)
        {
            DoorTypeMirrorEnsureResult result = new DoorTypeMirrorEnsureResult();

            if (doc == null || sourceSymbol == null || sourceSymbol.Family == null)
                return result;

            string sourceTypeName = sourceSymbol.Name ?? "";
            DoorOpeningMarker marker = GetDoorOpeningMarkerFromTypeName(sourceTypeName);

            if (marker == DoorOpeningMarker.None)
                return result;

            string leftName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.Left);
            string rightName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.Right);
            string rightAltName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.RightAlt);

            Family family = sourceSymbol.Family;
            List<FamilySymbol> familySymbols = GetAllSymbolsOfFamily(doc, family);

            FamilySymbol existingLeft = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, leftName, StringComparison.OrdinalIgnoreCase));

            FamilySymbol existingRight = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, rightName, StringComparison.OrdinalIgnoreCase));

            FamilySymbol existingRightAlt = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, rightAltName, StringComparison.OrdinalIgnoreCase));

            bool needCreateLeft = existingLeft == null;
            bool needCreateRight = existingRight == null && existingRightAlt == null;

            if (!needCreateLeft && !needCreateRight)
                return result;

            using (Transaction t = new Transaction(doc, "KPLN. Создание парных типов двери"))
            {
                t.Start();

                if (needCreateLeft)
                {
                    string newTypeName = leftName;
                    ElementType duplicated = sourceSymbol.Duplicate(newTypeName) as ElementType;
                    FamilySymbol newSymbol = duplicated as FamilySymbol;

                    if (newSymbol == null)
                        throw new Exception("Не удалось создать типоразмер '" + newTypeName + "'.");

                    Parameter p = FindDoorLeftOpeningParameter(newSymbol);
                    if (p == null)
                        throw new Exception("У нового типа '" + newTypeName + "' не найден параметр 'КП_О_Левое открывание'.");

                    if (!SetYesNoParameter(p, true))
                        throw new Exception("Не удалось включить параметр 'КП_О_Левое открывание' у типа '" + newTypeName + "'.");

                    result.HasMessage = true;
                    result.Message =
                        "Создан парный левый тип двери: '" + newTypeName + "'.";
                }

                if (needCreateRight)
                {
                    string newTypeName = rightName;
                    ElementType duplicated = sourceSymbol.Duplicate(newTypeName) as ElementType;
                    FamilySymbol newSymbol = duplicated as FamilySymbol;

                    if (newSymbol == null)
                        throw new Exception("Не удалось создать типоразмер '" + newTypeName + "'.");

                    Parameter p = FindDoorLeftOpeningParameter(newSymbol);
                    if (p == null)
                        throw new Exception("У нового типа '" + newTypeName + "' не найден параметр 'КП_О_Левое открывание'.");

                    if (!SetYesNoParameter(p, false))
                        throw new Exception("Не удалось выключить параметр 'КП_О_Левое открывание' у типа '" + newTypeName + "'.");

                    result.HasMessage = true;

                    if (string.IsNullOrWhiteSpace(result.Message))
                        result.Message = "Создан парный правый тип двери: '" + newTypeName + "'.";
                    else
                        result.Message += "\nСоздан парный правый тип двери: '" + newTypeName + "'.";
                }

                t.Commit();
            }

            return result;
        }

        private FamilySymbol ResolveDoorSymbolForPlacement(Document doc, FamilyInstance source2DDoor, FamilySymbol baseDoorSymbol, List<string> debugMessages)
        {
            if (doc == null || source2DDoor == null || baseDoorSymbol == null)
                return baseDoorSymbol;

            string baseTypeName = baseDoorSymbol.Name ?? "";
            DoorOpeningMarker baseMarker = GetDoorOpeningMarkerFromTypeName(baseTypeName);

            if (baseMarker == DoorOpeningMarker.None)
                return baseDoorSymbol;

            DoorTypeMirrorEnsureResult ensureResult = EnsureDoorMirrorTypeExists(doc, baseDoorSymbol);
            if (ensureResult != null && ensureResult.HasMessage && debugMessages != null)
                debugMessages.Add(ensureResult.Message);

            bool faceFlip;
            bool mirrored;
            if (!TryGetDoorOrientationFlags(source2DDoor, out faceFlip, out mirrored))
                return baseDoorSymbol;

            DoorOpeningMarker desiredMarker = ResolveDesiredDoorOpeningMarker(faceFlip, mirrored);
            if (desiredMarker == DoorOpeningMarker.None)
                return baseDoorSymbol;

            DoorOpeningMarker normalizedBaseMarker = NormalizeDoorOpeningMarkerForSelection(baseMarker);
            DoorOpeningMarker normalizedDesiredMarker = NormalizeDoorOpeningMarkerForSelection(desiredMarker);

            if (normalizedBaseMarker == normalizedDesiredMarker)
                return baseDoorSymbol;

            string desiredTypeName = ReplaceDoorOpeningMarker(baseTypeName, desiredMarker);

            FamilySymbol resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, desiredTypeName);
            if (resolved != null)
                return resolved;

            if (desiredMarker == DoorOpeningMarker.Right)
            {
                string rightAltName = ReplaceDoorOpeningMarker(baseTypeName, DoorOpeningMarker.RightAlt);
                resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, rightAltName);
                if (resolved != null)
                    return resolved;
            }

            if (desiredMarker == DoorOpeningMarker.RightAlt)
            {
                string rightName = ReplaceDoorOpeningMarker(baseTypeName, DoorOpeningMarker.Right);
                resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, rightName);
                if (resolved != null)
                    return resolved;
            }

            return baseDoorSymbol;
        }

        private static bool TryGetDoorOrientationFlags(FamilyInstance doorFi, out bool faceFlip, out bool mirrored)
        {
            faceFlip = false;
            mirrored = false;

            if (doorFi == null)
                return false;

            bool hasFaceFlip = TryGetYesNoParamFromElementOrType(doorFi, out faceFlip, "faceFlip", "FaceFlip", "Facing Flip", "FacingFlipped");
            bool hasMirrored = TryGetYesNoParamFromElementOrType(doorFi, out mirrored, "mirrored", "Mirrored");

            if (!hasFaceFlip)
            {
                try
                {
                    faceFlip = doorFi.FacingFlipped;
                    hasFaceFlip = true;
                }
                catch
                {
                }
            }

            if (!hasMirrored)
            {
                try
                {
                    mirrored = doorFi.Mirrored;
                    hasMirrored = true;
                }
                catch
                {
                }
            }

            return hasFaceFlip && hasMirrored;
        }

        private static bool TryGetYesNoParamFromElementOrType(Element e, out bool value, params string[] paramNames)
        {
            value = false;

            if (e == null || paramNames == null || paramNames.Length == 0)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null)
                {
                    if (p.StorageType == StorageType.Integer)
                    {
                        value = p.AsInteger() != 0;
                        return true;
                    }

                    if (p.StorageType == StorageType.String)
                    {
                        string s = p.AsString();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            s = s.Trim();
                            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                            {
                                value = true;
                                return true;
                            }

                            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                            {
                                value = false;
                                return true;
                            }
                        }
                    }
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null)
                    {
                        if (p.StorageType == StorageType.Integer)
                        {
                            value = p.AsInteger() != 0;
                            return true;
                        }

                        if (p.StorageType == StorageType.String)
                        {
                            string s = p.AsString();
                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                s = s.Trim();
                                if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                                {
                                    value = true;
                                    return true;
                                }

                                if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                                {
                                    value = false;
                                    return true;
                                }
                            }
                        }
                    }
                }
            }

            return false;
        }

        private static DoorOpeningMarker ResolveDesiredDoorOpeningMarker(bool faceFlip, bool mirrored)
        {
            if (!faceFlip && !mirrored)
                return DoorOpeningMarker.Left;

            if (faceFlip && mirrored)
                return DoorOpeningMarker.Left;

            if (faceFlip && !mirrored)
                return DoorOpeningMarker.Right;

            if (!faceFlip && mirrored)
                return DoorOpeningMarker.Right;

            return DoorOpeningMarker.None;
        }

        private static FamilySymbol FindFamilySymbolByTypeName(Document doc, Family family, string typeName)
        {
            if (doc == null || family == null || string.IsNullOrWhiteSpace(typeName))
                return null;

            List<FamilySymbol> symbols = GetAllSymbolsOfFamily(doc, family);

            return symbols.FirstOrDefault(x =>
                string.Equals(x.Name, typeName, StringComparison.OrdinalIgnoreCase));
        }

        private static DoorOpeningMarker NormalizeDoorOpeningMarkerForSelection(DoorOpeningMarker marker)
        {
            if (marker == DoorOpeningMarker.RightAlt)
                return DoorOpeningMarker.Right;

            return marker;
        }


        private static List<FamilyInstance> FindDoorSubComponentsRecursive(Document doc, FamilyInstance rootInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (doc == null || rootInstance == null)
                return result;

            CollectDoorSubComponentsRecursive(doc, rootInstance, result);
            return result;
        }

        private static void CollectDoorSubComponentsRecursive(Document doc, FamilyInstance current, List<FamilyInstance> result)
        {
            if (doc == null || current == null)
                return;

            ICollection<ElementId> subIds = current.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                string familyName = "";
                string typeName = "";
                string categoryName = "";

                if (subFi.Symbol != null)
                {
                    typeName = subFi.Symbol.Name ?? "";
                    if (subFi.Symbol.Family != null)
                        familyName = subFi.Symbol.Family.Name ?? "";
                }

                if (subFi.Category != null)
                    categoryName = subFi.Category.Name ?? "";

                bool isGenericModel =
                    string.Equals(categoryName, "Обобщенные модели", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(categoryName, "Обобщённые модели", StringComparison.OrdinalIgnoreCase);

                bool isDoorFamily =
                    string.Equals(familyName, "Дверь", StringComparison.OrdinalIgnoreCase);

                if (isGenericModel && isDoorFamily)
                    result.Add(subFi);

                CollectDoorSubComponentsRecursive(doc, subFi, result);
            }
        }

        private PreparedApartmentRooms PrepareRoomsForApartment(Document doc, FamilyInstance apartmentFi, List<string> debugMessages)
        {
            PreparedApartmentRooms result = new PreparedApartmentRooms();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
            if (roomInstances == null || roomInstances.Count == 0)
                return result;

            foreach (FamilyInstance roomFi in roomInstances)
            {
                if (roomFi == null)
                    continue;

                try
                {
                    string roomName = GetRoomCategoryLabel(roomFi);
                    if (string.IsNullOrWhiteSpace(roomName))
                        roomName = "Помещение";

                    double expectedAreaInternal = 0;
                    TryGetAreaParamFromElementOrType(roomFi, out expectedAreaInternal, "КП_Р_Площадь", "КП_Р_ПЛОЩАДЬ");

                    Transform roomTransform = roomFi.GetTransform();
                    if (roomTransform == null)
                        continue;

                    XYZ insertPointInProject = roomTransform.Origin;

                    result.Rooms.Add(new PreparedRoomPlacement
                    {
                        ApartmentId = apartmentFi.Id,
                        RoomName = roomName.Trim(),
                        InsertPoint = insertPointInProject,
                        ExpectedAreaInternal = expectedAreaInternal
                    });
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add(
                            "Ошибка подготовки помещения ID = " +
                            GetElementIdValue(roomFi.Id) + ": " + ex.Message);
                }
            }

            return result;
        }

        private static bool HasRoomAtPoint(Document doc, Level level, XYZ point)
        {
            if (doc == null || level == null || point == null)
                return false;

            List<Room> rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(x => x != null && x.LevelId == level.Id && x.Area > 0)
                .ToList();

            foreach (Room room in rooms)
            {
                try
                {
                    if (room.IsPointInRoom(point))
                        return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private int PlaceRoomsForApartment(Document doc, PreparedApartmentRooms apartmentRooms, Level roomLevel,
            List<RoomAreaMismatchInfo> roomAreaMismatches, List<DeletedRoomMismatchInfo> deletedRoomMismatches, List<Wall> createdWallsForApartment)
        {
            if (doc == null || apartmentRooms == null || roomLevel == null)
                return 0;

            if (apartmentRooms.Rooms == null || apartmentRooms.Rooms.Count == 0)
                return 0;

            int createdCount = 0;
            double areaToleranceSquareMeters = 0.1;

            foreach (PreparedRoomPlacement preparedRoom in apartmentRooms.Rooms)
            {
                if (preparedRoom == null || preparedRoom.InsertPoint == null)
                    continue;

                try
                {
                    XYZ roomPoint = new XYZ(
                        preparedRoom.InsertPoint.X,
                        preparedRoom.InsertPoint.Y,
                        roomLevel.Elevation + ConvertMmToInternal(100));

                    if (HasRoomAtPoint(doc, roomLevel, roomPoint))
                        continue;

                    UV roomUv = new UV(preparedRoom.InsertPoint.X, preparedRoom.InsertPoint.Y);

                    Room createdRoom = doc.Create.NewRoom(roomLevel, roomUv);
                    if (createdRoom == null)
                        continue;

                    Parameter roomNameParam = createdRoom.get_Parameter(BuiltInParameter.ROOM_NAME);
                    if (roomNameParam != null && !roomNameParam.IsReadOnly && !string.IsNullOrWhiteSpace(preparedRoom.RoomName))
                        roomNameParam.Set(preparedRoom.RoomName);

                    if (preparedRoom.ExpectedAreaInternal > 0)
                    {
                        double actualAreaInternal = createdRoom.Area;
                        double expectedAreaSquareMeters = ConvertInternalAreaToSquareMeters(preparedRoom.ExpectedAreaInternal);
                        double actualAreaSquareMeters = ConvertInternalAreaToSquareMeters(actualAreaInternal);

                        if (Math.Abs(expectedAreaSquareMeters - actualAreaSquareMeters) > areaToleranceSquareMeters)
                        {
                            if (roomAreaMismatches != null)
                            {
                                roomAreaMismatches.Add(new RoomAreaMismatchInfo
                                {
                                    ApartmentId = preparedRoom.ApartmentId,
                                    RoomName = preparedRoom.RoomName,
                                    ExpectedAreaInternal = preparedRoom.ExpectedAreaInternal,
                                    ActualAreaInternal = actualAreaInternal
                                });
                            }

                            if (deletedRoomMismatches != null)
                            {
                                deletedRoomMismatches.Add(new DeletedRoomMismatchInfo
                                {
                                    ApartmentId = preparedRoom.ApartmentId,
                                    RoomName = preparedRoom.RoomName,
                                    ExpectedAreaInternal = preparedRoom.ExpectedAreaInternal,
                                    ActualAreaInternal = actualAreaInternal
                                });
                            }

                            doc.Delete(createdRoom.Id);
                            DeleteIsolatedCreatedWalls(doc, createdWallsForApartment);
                            continue;
                        }
                    }

                    createdCount++;
                }
                catch
                {
                }
            }

            return createdCount;
        }

        #region ГЕОМЕТРИЯ ПОМЕЩЕНИЙ И ПОСТРОЕНИЕ ОСЕЙ

        private static CurveLoop BuildRoomLoopFromInstance(FamilyInstance roomFi)
        {
            if (roomFi == null)
                throw new ArgumentNullException("roomFi");

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                throw new Exception("Не удалось получить Transform для вложенного помещения.");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            XYZ p1 = tr.OfPoint(new XYZ(-halfW, -halfD, 0));
            XYZ p2 = tr.OfPoint(new XYZ(halfW, -halfD, 0));
            XYZ p3 = tr.OfPoint(new XYZ(halfW, halfD, 0));
            XYZ p4 = tr.OfPoint(new XYZ(-halfW, halfD, 0));

            List<Curve> profile = new List<Curve>();
            profile.Add(Line.CreateBound(p1, p2));
            profile.Add(Line.CreateBound(p2, p3));
            profile.Add(Line.CreateBound(p3, p4));
            profile.Add(Line.CreateBound(p4, p1));

            return CurveLoop.Create(profile);
        }

        private static double GetRequiredLengthParam(Element e, params string[] paramNames)
        {
            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            Element typeElem = null;
            if (e != null && e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                        return p.AsDouble();
                }
            }

            throw new Exception("Не найден параметр: " + string.Join(", ", paramNames));
        }

        private static List<Line> BuildClosedWallAxisLinesFromRooms(List<CurveLoop> roomLoops, double wallWidth, List<string> debugMessages)
        {
            List<Line> result = new List<Line>();
            double halfWidth = wallWidth / 2.0;

            foreach (CurveLoop roomLoop in roomLoops)
            {
                try
                {
                    List<Line> offsetLoopLines = BuildOffsetClosedLoop(roomLoop, halfWidth);
                    result.AddRange(offsetLoopLines);
                }
                catch (Exception ex)
                {
                    debugMessages.Add("Ошибка построения замкнутого контура стены: " + ex.Message);
                }
            }

            return MergeCollinearLines(result);
        }

        private static List<Line> BuildOffsetClosedLoop(CurveLoop loop, double offset)
        {
            const double tol = 1e-9;

            List<Line> edges = loop.Cast<Curve>()
                .Select(x => x as Line)
                .Where(x => x != null && x.Length > tol)
                .ToList();

            if (edges.Count < 3)
                throw new Exception("Контур помещения должен содержать минимум 3 линейных сегмента.");

            List<XYZ> vertices = ExtractOrderedVertices(edges);
            if (vertices.Count < 3)
                throw new Exception("Не удалось извлечь вершины контура помещения.");

            bool ccw = GetSignedAreaXY(vertices) > 0.0;

            List<OffsetLine2D> offsetLines = new List<OffsetLine2D>();

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                XYZ dir = Normalize2D(b - a);

                if (dir == null)
                    throw new Exception("Обнаружен нулевой сегмент в контуре.");

                XYZ outward = ccw ? new XYZ(dir.Y, -dir.X, 0) : new XYZ(-dir.Y, dir.X, 0);

                XYZ oa = new XYZ(a.X + outward.X * offset, a.Y + outward.Y * offset, a.Z);
                XYZ ob = new XYZ(b.X + outward.X * offset, b.Y + outward.Y * offset, b.Z);

                offsetLines.Add(new OffsetLine2D
                {
                    P0 = oa,
                    P1 = ob,
                    Dir = dir,
                    Z = a.Z
                });
            }

            List<XYZ> offsetVertices = new List<XYZ>();

            for (int i = 0; i < offsetLines.Count; i++)
            {
                OffsetLine2D prev = offsetLines[(i - 1 + offsetLines.Count) % offsetLines.Count];
                OffsetLine2D cur = offsetLines[i];

                XYZ intersection;
                if (!TryIntersectInfiniteLines2D(prev.P0, prev.Dir, cur.P0, cur.Dir, out intersection))
                    intersection = cur.P0;

                double z = 0.5 * (prev.Z + cur.Z);
                offsetVertices.Add(new XYZ(intersection.X, intersection.Y, z));
            }

            List<Line> result = new List<Line>();

            for (int i = 0; i < offsetVertices.Count; i++)
            {
                XYZ p0 = offsetVertices[i];
                XYZ p1 = offsetVertices[(i + 1) % offsetVertices.Count];

                if (p0.DistanceTo(p1) > tol)
                    result.Add(Line.CreateBound(p0, p1));
            }

            return result;
        }

        private static List<XYZ> ExtractOrderedVertices(List<Line> lines)
        {
            const double tol = 1e-6;
            List<XYZ> pts = new List<XYZ>();

            if (lines == null || lines.Count == 0)
                return pts;

            foreach (Line line in lines)
            {
                XYZ p = line.GetEndPoint(0);
                if (pts.Count == 0 || pts[pts.Count - 1].DistanceTo(p) > tol)
                    pts.Add(p);
            }

            if (pts.Count >= 2 && pts[0].DistanceTo(pts[pts.Count - 1]) < tol)
                pts.RemoveAt(pts.Count - 1);

            return pts;
        }

        private static double GetSignedAreaXY(List<XYZ> pts)
        {
            if (pts == null || pts.Count < 3)
                return 0.0;

            double area2 = 0.0;

            for (int i = 0; i < pts.Count; i++)
            {
                XYZ a = pts[i];
                XYZ b = pts[(i + 1) % pts.Count];
                area2 += (a.X * b.Y) - (b.X * a.Y);
            }

            return area2 * 0.5;
        }

        private static List<Line> MergeCollinearLines(List<Line> lines)
        {
            const double tol = 1e-6;
            List<GenericAxisLineData> data = new List<GenericAxisLineData>();

            foreach (Line line in lines)
            {
                if (line == null || line.Length <= tol)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                dir = CanonicalizeDirection(dir);
                XYZ normal = new XYZ(-dir.Y, dir.X, 0);
                double offset = Dot2D(p0, normal);
                double z = 0.5 * (p0.Z + p1.Z);
                double t0 = Dot2D(p0, dir);
                double t1 = Dot2D(p1, dir);
                double from = Math.Min(t0, t1);
                double to = Math.Max(t0, t1);

                data.Add(new GenericAxisLineData
                {
                    Dir = dir,
                    Normal = normal,
                    Offset = offset,
                    Z = z,
                    From = from,
                    To = to
                });
            }

            List<IGrouping<GenericAxisGroupKey, GenericAxisLineData>> groups = data
                .GroupBy(x => new GenericAxisGroupKey(x.Dir, x.Offset, x.Z, tol))
                .ToList();

            List<Line> result = new List<Line>();

            foreach (IGrouping<GenericAxisGroupKey, GenericAxisLineData> group in groups)
            {
                List<GenericAxisLineData> ordered = group.OrderBy(x => x.From).ThenBy(x => x.To).ToList();

                if (ordered.Count == 0)
                    continue;

                double curFrom = ordered[0].From;
                double curTo = ordered[0].To;

                for (int i = 1; i < ordered.Count; i++)
                {
                    GenericAxisLineData next = ordered[i];

                    if (next.From <= curTo + tol)
                    {
                        curTo = Math.Max(curTo, next.To);
                    }
                    else
                    {
                        result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
                        curFrom = next.From;
                        curTo = next.To;
                    }
                }

                result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
            }

            return result;
        }

        private static Line BuildGenericAxisLine(XYZ dir, double offset, double z, double from, double to)
        {
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            XYZ p0 = new XYZ(dir.X * from + normal.X * offset, dir.Y * from + normal.Y * offset, z);
            XYZ p1 = new XYZ(dir.X * to + normal.X * offset, dir.Y * to + normal.Y * offset, z);

            return Line.CreateBound(p0, p1);
        }

        private static List<Line> SnapNewLinesToExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            List<Line> result = new List<Line>();

            foreach (Line line in newLines)
            {
                if (line == null || line.Length < 1e-9)
                    continue;

                Line snapped = SnapSingleLineToExistingWalls(line, existingWalls, snapTol);
                if (snapped != null && snapped.Length > 1e-9)
                    result.Add(snapped);
            }

            return result;
        }

        private static Line SnapSingleLineToExistingWalls(Line line, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);

            if (dir == null)
                return line;

            XYZ newP0 = SnapEndpointToExistingWalls(p0, new XYZ(-dir.X, -dir.Y, 0), line, existingWalls, snapTol);
            XYZ newP1 = SnapEndpointToExistingWalls(p1, dir, line, existingWalls, snapTol);

            if (newP0.DistanceTo(newP1) < 1e-9)
                return null;

            return Line.CreateBound(newP0, newP1);
        }

        private static XYZ SnapEndpointToExistingWalls(XYZ endpoint, XYZ extensionDir, Line sourceLine, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            const double tol = 1e-9;

            XYZ bestPoint = endpoint;
            double bestDist = double.MaxValue;

            XYZ sourceP0 = sourceLine.GetEndPoint(0);
            XYZ sourceP1 = sourceLine.GetEndPoint(1);
            XYZ sourceDir = Normalize2D(sourceP1 - sourceP0);

            if (sourceDir == null)
                return endpoint;

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (Math.Abs(ex.Z - endpoint.Z) > snapTol)
                    continue;

                XYZ inter;
                if (TryIntersectInfiniteLines2D(endpoint, extensionDir, ex.P0, ex.Dir, out inter))
                {
                    if (PointOnSegment2D(inter, ex.P0, ex.P1, snapTol))
                    {
                        XYZ delta = inter - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, inter);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(inter.X, inter.Y, endpoint.Z);
                        }
                    }
                }

                if (AreCollinear2D(sourceP0, sourceP1, ex.P0, ex.P1, snapTol))
                {
                    XYZ[] candidates = new XYZ[] { ex.P0, ex.P1 };

                    foreach (XYZ c in candidates)
                    {
                        XYZ delta = c - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, c);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(c.X, c.Y, endpoint.Z);
                        }
                    }
                }
            }

            return bestPoint;
        }

        private static List<Line> RemoveSegmentsOverlappingExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls)
        {
            const double tol = 1e-6;
            List<Line> result = new List<Line>();

            foreach (Line newLine in newLines)
            {
                if (newLine == null || newLine.Length <= tol)
                    continue;

                List<Line> remaining = SubtractExistingWallsFromNewLine(newLine, existingWalls);

                foreach (Line part in remaining)
                {
                    if (part != null && part.Length > tol)
                        result.Add(part);
                }
            }

            return result;
        }

        private static List<Line> SubtractExistingWallsFromNewLine(Line newLine, List<ExistingWallLineInfo> existingWalls)
        {
            const double tol = 1e-6;

            XYZ p0 = newLine.GetEndPoint(0);
            XYZ p1 = newLine.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);

            if (dir == null)
                return new List<Line>();

            dir = CanonicalizeDirection(dir);
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            double newOffset = Dot2D(p0, normal);
            double newZ = 0.5 * (p0.Z + p1.Z);

            double newT0 = Dot2D(p0, dir);
            double newT1 = Dot2D(p1, dir);
            double newFrom = Math.Min(newT0, newT1);
            double newTo = Math.Max(newT0, newT1);

            List<Interval1D> cutIntervals = new List<Interval1D>();

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (ex == null)
                    continue;

                if (Math.Abs(ex.Z - newZ) > tol)
                    continue;

                XYZ exDir = CanonicalizeDirection(ex.Dir);
                if (exDir == null)
                    continue;

                if (Math.Abs(Cross2D(dir, exDir)) > tol)
                    continue;

                double exOffset = Dot2D(ex.P0, normal);
                if (Math.Abs(exOffset - newOffset) > tol)
                    continue;

                double exT0 = Dot2D(ex.P0, dir);
                double exT1 = Dot2D(ex.P1, dir);
                double exFrom = Math.Min(exT0, exT1);
                double exTo = Math.Max(exT0, exT1);

                double overlapFrom = Math.Max(newFrom, exFrom);
                double overlapTo = Math.Min(newTo, exTo);

                if (overlapTo - overlapFrom > tol)
                {
                    cutIntervals.Add(new Interval1D
                    {
                        From = overlapFrom,
                        To = overlapTo
                    });
                }
            }

            if (cutIntervals.Count == 0)
                return new List<Line> { newLine };

            List<Interval1D> mergedCuts = MergeIntervals(cutIntervals);
            List<Interval1D> remainingIntervals = SubtractIntervals(
                new Interval1D { From = newFrom, To = newTo },
                mergedCuts);

            List<Line> result = new List<Line>();

            foreach (Interval1D interval in remainingIntervals)
            {
                if (interval.To - interval.From <= tol)
                    continue;

                XYZ rp0 = new XYZ(dir.X * interval.From + normal.X * newOffset, dir.Y * interval.From + normal.Y * newOffset, newZ);
                XYZ rp1 = new XYZ(dir.X * interval.To + normal.X * newOffset, dir.Y * interval.To + normal.Y * newOffset, newZ);

                if (rp0.DistanceTo(rp1) > tol)
                    result.Add(Line.CreateBound(rp0, rp1));
            }

            return result;
        }

        private static List<Interval1D> MergeIntervals(List<Interval1D> intervals)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (intervals == null || intervals.Count == 0)
                return result;

            List<Interval1D> ordered = intervals
                .Where(x => x != null && x.To - x.From > tol)
                .OrderBy(x => x.From)
                .ThenBy(x => x.To)
                .ToList();

            if (ordered.Count == 0)
                return result;

            double curFrom = ordered[0].From;
            double curTo = ordered[0].To;

            for (int i = 1; i < ordered.Count; i++)
            {
                Interval1D next = ordered[i];

                if (next.From <= curTo + tol)
                {
                    curTo = Math.Max(curTo, next.To);
                }
                else
                {
                    result.Add(new Interval1D { From = curFrom, To = curTo });
                    curFrom = next.From;
                    curTo = next.To;
                }
            }

            result.Add(new Interval1D { From = curFrom, To = curTo });
            return result;
        }

        private static List<Interval1D> SubtractIntervals(Interval1D source, List<Interval1D> cuts)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (source == null || source.To - source.From <= tol)
                return result;

            if (cuts == null || cuts.Count == 0)
            {
                result.Add(source);
                return result;
            }

            double cursor = source.From;

            foreach (Interval1D cut in cuts)
            {
                if (cut == null || cut.To - cut.From <= tol)
                    continue;

                if (cut.To <= source.From + tol)
                    continue;

                if (cut.From >= source.To - tol)
                    break;

                double cutFrom = Math.Max(source.From, cut.From);
                double cutTo = Math.Min(source.To, cut.To);

                if (cutFrom > cursor + tol)
                {
                    result.Add(new Interval1D
                    {
                        From = cursor,
                        To = cutFrom
                    });
                }

                cursor = Math.Max(cursor, cutTo);
            }

            if (cursor < source.To - tol)
            {
                result.Add(new Interval1D
                {
                    From = cursor,
                    To = source.To
                });
            }

            return result;
        }

        private static bool IntersectsExistingWalls(Line candidate, double candidateThicknessInternal, List<ExistingWallLineInfo> existingWalls, double tol)
        {
            if (candidate == null || existingWalls == null || existingWalls.Count == 0)
                return false;

            XYZ a0 = candidate.GetEndPoint(0);
            XYZ a1 = candidate.GetEndPoint(1);
            XYZ aDir = Normalize2D(a1 - a0);

            if (aDir == null)
                return false;

            double aZ = 0.5 * (a0.Z + a1.Z);

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (ex == null)
                    continue;

                if (Math.Abs(ex.Z - aZ) > tol)
                    continue;

                if (HasForbiddenAxisIntersection2D(a0, a1, ex.P0, ex.P1, tol))
                    return true;

                if (HasParallelWallBodyOverlap(candidate, candidateThicknessInternal, ex, tol))
                    return true;
            }

            return false;
        }

        private static bool HasForbiddenAxisIntersection2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            if (a0 == null || a1 == null || b0 == null || b1 == null)
                return false;

            if (AreCollinear2D(a0, a1, b0, b1, tol))
            {
                XYZ dir = Normalize2D(a1 - a0);
                if (dir == null)
                    return false;

                double aT0 = Dot2D(a0, dir);
                double aT1 = Dot2D(a1, dir);
                double bT0 = Dot2D(b0, dir);
                double bT1 = Dot2D(b1, dir);

                double aFrom = Math.Min(aT0, aT1);
                double aTo = Math.Max(aT0, aT1);
                double bFrom = Math.Min(bT0, bT1);
                double bTo = Math.Max(bT0, bT1);

                double overlapFrom = Math.Max(aFrom, bFrom);
                double overlapTo = Math.Min(aTo, bTo);
                double overlapLen = overlapTo - overlapFrom;

                return overlapLen > tol * 10.0;
            }

            XYZ intersection;
            if (!TryIntersectSegments2D(a0, a1, b0, b1, out intersection, tol))
                return false;

            bool isEndpointOfA =
                Distance2D(intersection, a0) <= tol ||
                Distance2D(intersection, a1) <= tol;

            bool isEndpointOfB =
                Distance2D(intersection, b0) <= tol ||
                Distance2D(intersection, b1) <= tol;

            if (isEndpointOfA || isEndpointOfB)
                return false;

            return true;
        }

        private static bool HasParallelWallBodyOverlap(Line candidate, double candidateThicknessInternal, ExistingWallLineInfo existingWall, double tol)
        {
            if (candidate == null || existingWall == null)
                return false;

            XYZ a0 = candidate.GetEndPoint(0);
            XYZ a1 = candidate.GetEndPoint(1);

            XYZ aDir = Normalize2D(a1 - a0);
            XYZ bDir = CanonicalizeDirection(existingWall.Dir);

            if (aDir == null || bDir == null)
                return false;

            aDir = CanonicalizeDirection(aDir);

            if (Math.Abs(Cross2D(aDir, bDir)) > tol)
                return false;

            XYZ normal = new XYZ(-aDir.Y, aDir.X, 0);

            double aOffset = Dot2D(a0, normal);
            double bOffset = Dot2D(existingWall.P0, normal);
            double axisDistance = Math.Abs(aOffset - bOffset);

            double existingThicknessInternal = existingWall.ThicknessInternal > 0
                ? existingWall.ThicknessInternal
                : 0;

            double halfSumThickness = (candidateThicknessInternal + existingThicknessInternal) / 2.0;

            double penetrationDepth = halfSumThickness - axisDistance;
            if (penetrationDepth <= tol)
                return false;

            double aT0 = Dot2D(a0, aDir);
            double aT1 = Dot2D(a1, aDir);
            double aFrom = Math.Min(aT0, aT1);
            double aTo = Math.Max(aT0, aT1);

            double bT0 = Dot2D(existingWall.P0, aDir);
            double bT1 = Dot2D(existingWall.P1, aDir);
            double bFrom = Math.Min(bT0, bT1);
            double bTo = Math.Max(bT0, bT1);

            double overlapFrom = Math.Max(aFrom, bFrom);
            double overlapTo = Math.Min(aTo, bTo);

            return (overlapTo - overlapFrom) > tol;
        }

        private static XYZ Normalize2D(XYZ v)
        {
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-12)
                return null;

            return new XYZ(v.X / len, v.Y / len, 0);
        }

        private static XYZ CanonicalizeDirection(XYZ dir)
        {
            if (dir == null)
                return null;

            if (dir.X < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            if (Math.Abs(dir.X) < 1e-9 && dir.Y < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            return new XYZ(dir.X, dir.Y, 0);
        }

        private static double Dot2D(XYZ a, XYZ b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static bool PointOnSegment2D(XYZ p, XYZ a, XYZ b, double tol)
        {
            double ab = Distance2D(a, b);
            double ap = Distance2D(a, p);
            double pb = Distance2D(p, b);
            return Math.Abs((ap + pb) - ab) <= tol;
        }

        private static bool AreCollinear2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            if (Math.Abs(Cross2D(ad, bd)) > tol)
                return false;

            if (Math.Abs(Cross2D(b0 - a0, ad)) > tol)
                return false;

            return true;
        }

        private static bool TryIntersectInfiniteLines2D(XYZ p1, XYZ d1, XYZ p2, XYZ d2, out XYZ intersection)
        {
            const double tol = 1e-12;
            intersection = null;

            double cross = Cross2D(d1, d2);
            if (Math.Abs(cross) < tol)
                return false;

            XYZ delta = p2 - p1;
            double t = Cross2D(delta, d2) / cross;

            intersection = new XYZ(
                p1.X + d1.X * t,
                p1.Y + d1.Y * t,
                0);

            return true;
        }

        private static bool TryIntersectSegments2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, out XYZ intersection, double tol)
        {
            intersection = null;

            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            XYZ inter;
            if (!TryIntersectInfiniteLines2D(a0, ad, b0, bd, out inter))
                return false;

            if (!PointOnSegment2D(inter, a0, a1, tol))
                return false;

            if (!PointOnSegment2D(inter, b0, b1, tol))
                return false;

            intersection = inter;
            return true;
        }

        #endregion
    }
}
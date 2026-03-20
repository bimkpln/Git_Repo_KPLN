using Autodesk.Revit.Attributes;
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
using System.Windows.Interop;
using System.Windows.Threading;

namespace KPLN_Tools.ExecutableCommand
{
    /// <summary>
    /// Команда открытия основного окна Apartment Manager.
    /// 1. Проверить, не открыто ли уже окно.
    /// 2. Если окно открыто — просто показать и активировать его.
    /// 3. Если окна нет — создать контроллер, окно и связать их между собой.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandApartmentManagerShow : IExternalCommand
    {
        private static ApartmentManagerWindow _window;
        private static ApartmentManagerExternalController _controller;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;

                // Если окно уже существует — просто показываем его поверх.
                if (_window != null)
                {
                    ShowAndActivateExistingWindow();
                    return Result.Succeeded;
                }

                _controller = new ApartmentManagerExternalController();
                _window = new ApartmentManagerWindow(DBWorkerService.CurrentDBUserSubDepartment.Id, _controller);
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

        /// <summary>
        /// Показать уже существующее окно и вернуть ему фокус. Используется вместо создания нового экземпляра окна.
        /// </summary>
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

        /// <summary>
        /// Очистка ссылок после закрытия окна.
        /// </summary>
        private static void OnWindowClosed(object sender, EventArgs e)
        {
            if (_controller != null)
                _controller.DetachWindow();

            _controller = null;
            _window = null;
        }
    }

    /// <summary>
    /// Контроллер-посредник между WPF-окном и ExternalEvent.
    /// </summary>
    internal class ApartmentManagerExternalController : IApartmentManagerExternalController
    {
        private readonly ApartmentManagerExternalHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public ApartmentManagerExternalController()
        {
            _handler = new ApartmentManagerExternalHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        /// <summary>
        /// Привязка окна к хендлеру.
        /// </summary>
        public void AttachWindow(ApartmentManagerWindow window)
        {
            _handler.AttachWindow(window);
        }

        /// <summary>
        /// Отвязка окна от хендлера.
        /// </summary>
        public void DetachWindow()
        {
            _handler.DetachWindow();
        }

        /// <summary>
        /// Запрос на размещение 2D-семейства квартиры.
        /// </summary>
        public void RequestPlaceApartment(int apartmentId)
        {
            _handler.PreparePlaceApartment(apartmentId);
            _externalEvent.Raise();
        }

        /// <summary>
        /// Запрос на конвертацию 2D-квартир в 3D-стены.
        /// </summary>
        public void RequestConvertTo3D(ApartmentPresetData presetData)
        {
            _handler.PrepareConvertTo3D(presetData);
            _externalEvent.Raise();
        }

        /// <summary>
        /// Запрос на открытие окна пресетов.
        /// </summary>
        public void RequestOpenApartmentPresets(ApartmentPresetData presetData)
        {
            _handler.PrepareOpenApartmentPresets(presetData);
            _externalEvent.Raise();
        }
    }

    /// <summary>
    /// Хендлер ExternalEvent.
    /// </summary>
    internal class ApartmentManagerExternalHandler : IExternalEventHandler
    {
        /// <summary>
        /// Тип запроса, который должен быть выполнен при очередном Execute().
        /// </summary>
        private enum RequestType
        {
            None,
            PlaceApartment,
            ConvertTo3D,
            OpenApartmentPresets
        }

        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string ApartmentInstanceMarker = "[KPLN_APT_INSTANCE]";

        private RequestType _requestType = RequestType.None;
        private int _requestedApartmentId; // ID квартиры для размещения. Используется только если _requestType == PlaceApartment.
        private ApartmentPresetData _requestedPresetData; // Набор пресетных данных для операций конвертации / открытия окна пресетов.
        private ApartmentManagerWindow _window;

        public string GetName()
        {
            return "KPLN Apartment Manager External Event";
        }

        /// <summary>
        /// Привязка окна к хендлеру.
        /// </summary>
        public void AttachWindow(ApartmentManagerWindow window)
        {
            _window = window;
        }

        /// <summary>
        /// Отвязка окна от хендлера.
        /// </summary>
        public void DetachWindow()
        {
            _window = null;
        }

        /// <summary>
        /// Подготовить запрос на размещение квартиры.
        /// </summary>
        public void PreparePlaceApartment(int apartmentId)
        {
            _requestType = RequestType.PlaceApartment;
            _requestedApartmentId = apartmentId;
            _requestedPresetData = null;
        }

        /// <summary>
        /// Подготовить запрос на построение 3D-стен.
        /// </summary>
        public void PrepareConvertTo3D(ApartmentPresetData presetData)
        {
            _requestType = RequestType.ConvertTo3D;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        /// <summary>
        /// Подготовить запрос на открытие окна пресетов.
        /// </summary>
        public void PrepareOpenApartmentPresets(ApartmentPresetData presetData)
        {
            _requestType = RequestType.OpenApartmentPresets;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        /// <summary>
        /// Возвращает фокус окну Apartment Manager после завершения внешнего события.
        /// </summary>
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

        ///// Вспомогательные DTO и служебные классы
        /// <summary>
        /// Подготовленный набор осей стен для одной квартиры.
        /// </summary>
        private class PreparedApartmentWalls
        {
            public ElementId ApartmentId { get; set; }
            public WallType WallType { get; set; }
            public int ThicknessMm { get; set; }
            public List<Line> AxisLines { get; set; }
        }

        /// <summary>
        /// Смещённая 2D-линия для построения offset-контура.
        /// </summary>
        private class OffsetLine2D
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        /// <summary>
        /// Геометрическое представление существующей стены как осевой линии.
        /// </summary>
        private class ExistingWallLineInfo
        {
            public ElementId WallId { get; set; }
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        /// <summary>
        /// Одномерный интервал.
        /// Используется при вычитании перекрывающихся участков линий.
        /// </summary>
        private class Interval1D
        {
            public double From { get; set; }
            public double To { get; set; }
        }

        /// <summary>
        /// Нормализованное представление осевой линии для группировки.
        /// </summary>
        private class GenericAxisLineData
        {
            public XYZ Dir { get; set; }
            public XYZ Normal { get; set; }
            public double Offset { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
        }

        /// <summary>
        /// Ключ группировки коллинеарных линий. Содержит нормализованное направление, смещение и высоту Z.
        /// </summary>
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

        /// <summary>
        /// Округлить значение до заданного допуска. Используется при формировании ключей группировки линий.
        /// </summary>
        private static double RoundTol(double value, double tol)
        {
            return Math.Round(value / tol) * tol;
        }

        /// <summary>
        /// Унифицированный доступ к числовому значению ElementId
        /// для разных версий Revit API.
        /// </summary>
        private static long GetElementIdValue(ElementId id)
        {
#if Revit2024 || Debug2024
            return id.Value;
#else
            return id.IntegerValue;
#endif
        }

        /// <summary>
        /// Перевести миллиметры во внутренние единицы Revit.
        /// </summary>
        private static double ConvertMmToInternal(int valueMm)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        /// <summary>
        /// Перевести внутренние единицы Revit в миллиметры.
        /// </summary>
        private static double ConvertInternalToMm(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

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
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Пользователь отменил действие в Revit (например, PickPoint).
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

        /// <summary>
        /// Открыть модальное окно пресетов поверх главного окна Apartment Manager.
        /// </summary>
        private void ExecuteOpenApartmentPresets(UIDocument uidoc, Document doc, ApartmentPresetData currentPreset)
        {
            ViewPlan activeFloorPlan = doc.ActiveView as ViewPlan;
            if (activeFloorPlan == null || activeFloorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("Apartment Manager", "Перед открытием преднастроек откройте план этажа.");
                return;
            }

            ApartmentPresetWindowContext context = BuildPresetWindowContext(doc, activeFloorPlan);

            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                var wnd = new ApartmentPresetsWindow(
                    currentPreset != null ? currentPreset.Clone() : new ApartmentPresetData(),
                    context);

                wnd.Owner = _window;

                bool? res = wnd.ShowDialog();
                if (res == true && wnd.ResultPresetData != null)
                    _window.SetApartmentPresetData(wnd.ResultPresetData.Clone());
            }));
        }

        /// <summary>
        /// Собрать контекст для окна пресетов:
        /// - список планов;
        /// - нижние ограничения;
        /// - доступные толщины;
        /// - доступные типы стен под толщины;
        /// - категории помещений.
        /// </summary>
        private ApartmentPresetWindowContext BuildPresetWindowContext(Document doc, ViewPlan activeFloorPlan)
        {
            ApartmentPresetWindowContext context = new ApartmentPresetWindowContext();

            List<ViewPlan> plans = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().Where(x => !x.IsTemplate && x.ViewType == ViewType.FloorPlan && x.GenLevel != null)
                .OrderBy(x => x.Id != activeFloorPlan.Id).ThenBy(x => x.Name).ToList();

            foreach (ViewPlan plan in plans)
            {
                ApartmentPlanPresetOption option = new ApartmentPlanPresetOption();
                option.PlanName = plan.Name;
                option.LowerConstraintText = BuildLowerConstraintTextForPlan(doc, plan);
                option.UpperConstraintText = "Неприсоединённая";
                option.WallThicknesses = BuildWallThicknessesForPlan(doc, plan);
                option.WallTypeOptionsByThickness = BuildWallTypeOptionsByThicknessForPlan(doc, plan);
                option.RoomCategories = BuildRoomCategoriesForPlan(doc, plan);

                if (option.RoomCategories == null || option.RoomCategories.Count == 0)
                    option.RoomCategories = new List<string> { "Помещение" };

                context.Plans.Add(option);
            }

            return context;
        }

        /// <summary>
        /// Текст для окна присетов по привязки к нижнему уровню стен
        /// </summary>
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

        /// <summary>
        /// Собрать список уникальных толщин стен квартир на плане.
        /// </summary>
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

        /// <summary>
        /// Для каждой толщины стены собрать список доступных типов стен из проекта.
        /// </summary>
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

        /// <summary>
        /// Попытка получить толщину типа стены в миллиметрах. Сначала пробуем параметр "Толщина", затем стандартный Width.
        /// </summary>
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
            catch { }

            return false;
        }

        /// <summary>
        /// Собрать список категорий помещений по вложенным "комнатам" квартир на плане.
        /// </summary>
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

        /// <summary>
        /// Получить все квартиры Apartment Manager на заданном уровне.
        /// </summary>
        private static List<FamilyInstance> GetPlacedApartmentInstancesOnLevel(Document doc, Level level)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>();

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

        /// <summary>
        /// Безопасно попытаться получить толщину стены квартиры.
        /// </summary>
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

        /// <summary>
        /// Попытка получить человекочитаемую категорию помещения.
        /// </summary>
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

        /// <summary>
        /// Предварительная валидация пресета перед построением 3D.
        /// </summary>
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
                DoorsByRoomCategory = new Dictionary<string, string>()
            };

            List<string> problems = new List<string>();

            // Проверка высоты стены.
            if (effectivePreset.WallHeight <= 0)
                problems.Add("Не указана неприсоединённая высота стены.");

            // Поиск плана, на котором будет выполняться построение.
            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
            if (targetPlan == null)
            {
                problems.Add("Не удалось определить план для построения стен.");
            }
            else
            {
                // Собираем толщины стен из уже размещённых 2D-квартир.
                List<int> requiredThicknesses = BuildWallThicknessesForPlan(doc, targetPlan);

                if (requiredThicknesses.Count == 0)
                {
                    problems.Add("На выбранном плане не найдено ни одного размещённого 2D-семейства квартиры.");
                }
                else
                {
                    // Для каждой найденной толщины должен быть выбран корректный тип стены.
                    foreach (int thickness in requiredThicknesses.OrderBy(x => x))
                    {
                        string selectedWallType = GetSelectedWallTypeNameForThickness(effectivePreset, thickness);

                        if (string.IsNullOrWhiteSpace(selectedWallType) ||
                            string.Equals(selectedWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        {
                            problems.Add("Не выбран тип стены для толщины " + thickness + " мм.");
                            continue;
                        }

                        WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallType, thickness);
                        if (matchedWallType == null)
                        {
                            problems.Add("Для толщины " + thickness + " мм выбран тип стены '" + selectedWallType + "', но такой тип не найден в проекте.");
                        }
                    }
                }
            }

            if (problems.Count > 0)
            {
                validationMessage = "Невозможно выполнить построение 3D.\n" + "Необходимо заполнить/исправить:\n- " + string.Join("\n- ", problems);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Найти целевой план: по имени из пресета, иначе активный план.
        /// </summary>
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

        /// <summary>
        /// Получить выбранное имя типа стены для конкретной толщины из пресета.
        /// </summary>
        private static string GetSelectedWallTypeNameForThickness(ApartmentPresetData preset, int thicknessMm)
        {
            if (preset == null || preset.WallTypeByThickness == null || preset.WallTypeByThickness.Count == 0)
                return null;

            string value;
            if (preset.WallTypeByThickness.TryGetValue(thicknessMm, out value))
                return value;

            return null;
        }

        /// <summary>
        /// Найти тип стены строго по имени и толщине.
        /// </summary>
        private static WallType FindWallTypeByExactSelectionAndThickness(Document doc, string selectedWallTypeName, int thicknessMm)
        {
            if (doc == null || string.IsNullOrWhiteSpace(selectedWallTypeName))
                return null;

            List<WallType> allWallTypes = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();

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

        ////////////////////////////////////////////
        ////////////////////////////////////////////  ПОСТАНОВКА 2D-СЕМЕЙСТВА
        //////////////////////////////////////////// 
        /// <summary>
        /// Размещения 2D-семейства квартиры на активном плане.
        /// </summary>
        private void ExecutePlaceApartment(UIDocument uidoc, Document doc, int id)
        {
            ViewPlan floorPlan = doc.ActiveView as ViewPlan;
            if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("Предупреждение", "Откройте план этажа перед размещением семейства.");
                return;
            }

            // По ID квартиры получаем путь к семейству из БД.
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

            XYZ insertPoint = uidoc.Selection.PickPoint("Укажите точку вставки семейства");
            using (Transaction t = new Transaction(doc, "Разместить семейство квартиры"))
            {
                t.Start();

                Family family = LoadOrFindFamily(doc, familyPath);
                if (family == null)
                    throw new Exception("Не удалось загрузить или найти семейство в проекте.");

                // Берём первый попавшийся тип семейства.
                FamilySymbol symbol = GetFirstFamilySymbol(doc, family);
                if (symbol == null)
                    throw new Exception("У семейства не найдено ни одного типоразмера.");

                // Перед размещением тип должен быть активирован.
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }

                FamilyInstance placedInstance = PlaceFamilyInstance(doc, floorPlan, symbol, insertPoint);
                if (placedInstance == null)
                    throw new Exception("Не удалось разместить экземпляр семейства.");

                AppendComment(placedInstance, ApartmentInstanceMarker);

                t.Commit();
            }
        }

        /// <summary>
        /// Получить путь к семейству квартиры из SQLite-базы по ID.
        /// </summary>
        private static string GetFamilyPathById(int id)
        {
            if (!File.Exists(DbPath))
                throw new FileNotFoundException("Не найдена база данных", DbPath);

            using (var con = OpenConnection(DbPath, true))
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT FPATH FROM Main WHERE ID = @id LIMIT 1;";
                cmd.Parameters.AddWithValue("@id", id);

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    return null;

                return result.ToString().Trim();
            }
        }

        /// <summary>
        /// Открыть SQLite-подключение.
        /// </summary>
        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? ("Data Source=" + dbPath + ";Version=3;Read Only=True;")
                : ("Data Source=" + dbPath + ";Version=3;");

            SQLiteConnection con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

        /// <summary>
        /// Найти семейство в документе или загрузить его из файла.
        /// </summary>
        private static Family LoadOrFindFamily(Document doc, string familyPath)
        {
            string familyName = Path.GetFileNameWithoutExtension(familyPath);

            Family existingFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family)).Cast<Family>().FirstOrDefault(f => string.Equals(f.Name, familyName, StringComparison.OrdinalIgnoreCase));

            if (existingFamily != null)
                return existingFamily;

            Family loadedFamily;
            bool loaded = doc.LoadFamily(familyPath, out loadedFamily);

            if (loaded && loadedFamily != null)
                return loadedFamily;

            existingFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family)).Cast<Family>().FirstOrDefault(f => string.Equals(f.Name, familyName, StringComparison.OrdinalIgnoreCase));

            return existingFamily;
        }

        /// <summary>
        /// Получить первый типоразмер семейства
        /// </summary>
        private static FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
            if (symbolIds == null || symbolIds.Count == 0)
                return null;

            ElementId firstId = symbolIds.First();
            return doc.GetElement(firstId) as FamilySymbol;
        }

        /// <summary>
        /// Разместить семейство на плане с учётом его типа размещения.
        /// </summary>
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

        /// <summary>
        /// Добавить текст в комментарий элемента
        /// </summary>
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

        //////////////////////////////////////////// 
        //////////////////////////////////////////// ПРЕВРАЩЕНИЕ 2D-СЕМЕЙСТВ В СТЕНЫ И ДВЕРИ
        //////////////////////////////////////////// 
        /// <summary>
        /// Построения 3D-стен и дверей по размещённым 2D-квартирам.
        /// </summary>
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
                DoorsByRoomCategory = new Dictionary<string, string>()
            };

            ViewPlan targetPlan = FindTargetFloorPlan(doc, effectivePreset.SelectedPlanName);
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
                TaskDialog.Show("Apartment Manager", "На выбранном плане не найдено ранее размещённых экземпляров квартир.");
                return;
            }

            // Коллекции для отладки и отчёта.
            List<string> debugMessages = new List<string>();
            List<PreparedApartmentWalls> preparedApartments = new List<PreparedApartmentWalls>();
        
            double connectTol = ConvertMmToInternal(150); // Допуск привязки новых осей к существующим стенам.
            List<ExistingWallLineInfo> existingWalls = GetExistingWallLinesOnLevel(doc, targetPlan.GenLevel.Id); // Существующие стены на уровне плана

            foreach (FamilyInstance apartmentFi in apartmentInstances)
            {
                try
                {
                    // Определяем толщину стен 2D-квартиры из параметра семейства.
                    double apartmentWallThicknessInternal = GetApartmentWallThickness(apartmentFi);
                    int apartmentWallThicknessMm = (int)Math.Round(ConvertInternalToMm(apartmentWallThicknessInternal));

                    if (apartmentWallThicknessMm <= 0)
                    {
                        debugMessages.Add("У квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " параметр 'Стены_Толщина' имеет некорректное значение.");
                        continue;
                    }

                    // Находим выбранный в пресете тип стены для этой толщины.
                    string selectedWallTypeName = GetSelectedWallTypeNameForThickness(effectivePreset, apartmentWallThicknessMm);
                    // Проверяем, что такой тип стены реально существует в проекте.
                    WallType matchedWallType = FindWallTypeByExactSelectionAndThickness(doc, selectedWallTypeName, apartmentWallThicknessMm);
                    if (matchedWallType == null)
                    {
                        debugMessages.Add("Для квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " не найден тип стены '" + selectedWallTypeName + "' с толщиной " + apartmentWallThicknessMm + " мм.");
                        continue;
                    }

                    // Ищем вложенные помещения внутри семейства квартиры.
                    List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
                    if (roomInstances.Count == 0)
                    {
                        debugMessages.Add("Не найдены вложенные экземпляры 'Помещение' у экземпляра ID = " + GetElementIdValue(apartmentFi.Id));
                        continue;
                    }

                    // Для каждого вложенного помещения строим замкнутый контур.
                    List<CurveLoop> apartmentRoomLoops = new List<CurveLoop>();

                    foreach (FamilyInstance roomFi in roomInstances)
                    {
                        try
                        {
                            CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi);
                            apartmentRoomLoops.Add(roomLoop);
                        }
                        catch (Exception exRoom)
                        {
                            debugMessages.Add("Ошибка обработки вложенного помещения ID = " + GetElementIdValue(roomFi.Id) + ": " + exRoom.Message);
                        }
                    }

                    if (apartmentRoomLoops.Count == 0)
                        continue;

                    // По контурам помещений строим осевые линии будущих стен.
                    List<Line> wallAxisLines = BuildClosedWallAxisLinesFromRooms(apartmentRoomLoops, apartmentWallThicknessInternal, debugMessages);
                    if (wallAxisLines.Count == 0)
                    {
                        debugMessages.Add("Для квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " не удалось вычислить оси стен.");
                        continue;
                    }

                    // Дорабатываем оси:
                    // 1. Привязываем к существующим стенам.
                    // 2. Склеиваем коллинеарные линии.
                    // 3. Удаляем участки, перекрытые существующими стенами.
                    List<Line> preparedAxisLines = SnapNewLinesToExistingWalls(wallAxisLines, existingWalls, connectTol);
                    preparedAxisLines = MergeCollinearLines(preparedAxisLines);
                    preparedAxisLines = RemoveSegmentsOverlappingExistingWalls(preparedAxisLines, existingWalls);
                    preparedAxisLines = MergeCollinearLines(preparedAxisLines);

                    if (preparedAxisLines.Count == 0)
                    {
                        debugMessages.Add("Для квартиры ID = " + GetElementIdValue(apartmentFi.Id) + " все вычисленные оси перекрыты существующими стенами.");
                        continue;
                    }

                    // Сохраняем подготовленные оси на дальнейшее создание стен.
                    preparedApartments.Add(new PreparedApartmentWalls
                    {
                        ApartmentId = apartmentFi.Id,
                        WallType = matchedWallType,
                        ThicknessMm = apartmentWallThicknessMm,
                        AxisLines = preparedAxisLines
                    });

                }
                catch (Exception exApartment)
                {
                    debugMessages.Add("Ошибка обработки квартиры ID = " + GetElementIdValue(apartmentFi.Id) + ": " + exApartment.Message);
                }
            }

            if (preparedApartments.Count == 0)
            {
                string debugText = debugMessages.Count > 0 ? "\n\n" + string.Join("\n", debugMessages) : "";
                TaskDialog.Show("Apartment Manager", "Не удалось подготовить ни одной стены." + debugText);
                return;
            }

            double baseOffsetInternal = ConvertMmToInternal(effectivePreset.BaseOffset);
            double wallHeightInternal = ConvertMmToInternal(effectivePreset.WallHeight > 0 ? effectivePreset.WallHeight : 3000);


            // Непосредственное создание стен.
            using (Transaction t = new Transaction(doc, "KPLN - Построение стен по помещениям"))
            {
                t.Start();

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null || apartmentWalls.WallType == null || apartmentWalls.AxisLines == null)
                        continue;

                    foreach (Line axis in apartmentWalls.AxisLines)
                    {
                        if (axis == null || axis.Length < 1e-6)
                            continue;

                        Wall wall = Wall.Create(doc, axis, apartmentWalls.WallType.Id, baseLevel.Id, wallHeightInternal, 0, false, false);

                        // Применяем верх/низ/смещение/высоту стены.
                        ApplyWallPresetParameters(wall, baseLevel, topLevel, baseOffsetInternal, wallHeightInternal);
                    }
                }

                t.Commit();
            }

            // Формируем итоговый отчёт.
            string report = "Построение завершено.\n\n" + "План: " + targetPlan.Name + "\n" + "Экземпляров квартир найдено: " + apartmentInstances.Count + "\n" + "Экземпляров квартир обработано: " + preparedApartments.Count + "\n";
            if (debugMessages.Count > 0)
                report += "\n\nЗамечания:\n" + string.Join("\n", debugMessages);

            TaskDialog.Show("ConvertTo3D", report);
        }

        /// <summary>
        /// Определить базовый уровень стены. Сначала по LowerConstraint пресета (или по уровню плана)
        /// </summary>
        private static Level ResolveBaseLevelForPreset(Document doc, ApartmentPresetData preset, ViewPlan targetPlan)
        {
            if (preset != null && !string.IsNullOrWhiteSpace(preset.LowerConstraint))
            {
                string[] parts = preset.LowerConstraint.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

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

        /// <summary>
        /// Определить верхний уровень стены. Если в пресете указано "Неприсоединённая" — верхний уровень не используется
        /// </summary>
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

        /// <summary>
        /// Найти уровень по имени
        /// </summary>
        private static Level FindLevelByName(Document doc, string levelName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(levelName))
                return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(x => string.Equals(x.Name, levelName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Собрать все ранее размещённые Apartment Manager квартиры для конкретного плана.
        /// </summary>
        private static List<FamilyInstance> GetPlacedApartmentInstancesForPlan(Document doc, ViewPlan plan)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            if (doc == null || plan == null)
                return result;

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>();

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

                // Если экземпляр явно принадлежит виду — считаем, что он на этом плане.
                if (fi.OwnerViewId != ElementId.InvalidElementId && fi.OwnerViewId == plan.Id)
                    belongsToPlan = true;

                // Если OwnerViewId не сработал — пытаемся сопоставить по уровню.
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

        /// <summary>
        /// Попытка определить уровень экземпляра семейства.
        /// </summary>
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

        /// <summary>
        /// Собрать осевые линии уже существующих стен на заданном уровне для стыковки и вычитания дублей.
        /// </summary>
        private static List<ExistingWallLineInfo> GetExistingWallLinesOnLevel(Document doc, ElementId levelId)
        {
            List<ExistingWallLineInfo> result = new List<ExistingWallLineInfo>();

            IEnumerable<Wall> walls = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>();

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

                result.Add(new ExistingWallLineInfo
                {
                    WallId = wall.Id,
                    P0 = p0,
                    P1 = p1,
                    Dir = dir,
                    Z = 0.5 * (p0.Z + p1.Z)
                });
            }

            return result;
        }

        /// <summary>
        /// Получить толщину стен у 2D-квартиры. В отличие от TryGetApartmentWallThickness, при неудаче кидает исключение.
        /// </summary>
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

        /// <summary>
        /// Найти вложенные семейства, которые трактуются как "Помещения". Поиск ведётся по имени семейства, имени типа или имени категории.
        /// </summary>
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

        /// <summary>
        /// Построить прямоугольный контур помещения по вложенному экземпляру.
        /// </summary>
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

        /// <summary>
        /// Получить обязательный параметр длины у экземпляра или его типа.
        /// </summary>
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

        /// <summary>
        /// Привязать все новые линии к ближайшим существующим стенам в пределах допуска.
        /// </summary>
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

        /// <summary>
        /// Привязать оба конца одной новой линии к существующим стенам.
        /// </summary>
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

        /// <summary>
        /// Привязать один конец линии к существующей стене. Ищется ближайшая допустимая точка либо по пересечению, либо по совпадающей коллинеарной геометрии.
        /// </summary>
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

        /// <summary>
        /// Расстояние между двумя точками в плоскости XY.
        /// </summary>
        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        /// <summary>
        /// Проверить, лежит ли точка на отрезке в 2D.
        /// </summary>
        private static bool PointOnSegment2D(XYZ p, XYZ a, XYZ b, double tol)
        {
            double ab = Distance2D(a, b);
            double ap = Distance2D(a, p);
            double pb = Distance2D(p, b);
            return Math.Abs((ap + pb) - ab) <= tol;
        }

        /// <summary>
        /// Проверить, лежат ли два отрезка на одной бесконечной прямой в 2D.
        /// </summary>
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

        /// <summary>
        /// Построить замкнутые оси стен по наборам контуров помещений. На каждый контур выполняется внешнее смещение на половину толщины стены.
        /// </summary>
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

        /// <summary>
        /// Построить смещённый наружу замкнутый контур из исходного прямолинейного контура.
        /// </summary>
        private static List<Line> BuildOffsetClosedLoop(CurveLoop loop, double offset)
        {
            const double tol = 1e-9;

            List<Line> edges = loop.Cast<Curve>().Select(x => x as Line).Where(x => x != null && x.Length > tol).ToList();

            if (edges.Count < 3)
                throw new Exception("Контур помещения должен содержать минимум 3 линейных сегмента.");

            List<XYZ> vertices = ExtractOrderedVertices(edges);
            if (vertices.Count < 3)
                throw new Exception("Не удалось извлечь вершины контура помещения.");

            // Определяем направление обхода контура (по/против часовой).
            bool ccw = GetSignedAreaXY(vertices) > 0.0;

            List<OffsetLine2D> offsetLines = new List<OffsetLine2D>();

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                XYZ dir = Normalize2D(b - a);

                if (dir == null)
                    throw new Exception("Обнаружен нулевой сегмент в контуре.");

                // Наружная нормаль зависит от ориентации обхода.
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

            // Пересекаем соседние смещённые прямые, чтобы получить вершины нового контура.
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

        /// <summary>
        /// Вытащить упорядоченные вершины из последовательности линий. Предполагается, что линии уже описывают замкнутый контур.
        /// </summary>
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

        /// <summary>
        /// Подписанная площадь полигона в XY. По знаку можно определить направление обхода.
        /// </summary>
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

        /// <summary>
        /// Нормализовать вектор в плоскости XY. Z принудительно обнуляется.
        /// </summary>
        private static XYZ Normalize2D(XYZ v)
        {
            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-12)
                return null;

            return new XYZ(v.X / len, v.Y / len, 0);
        }

        /// <summary>
        /// Пересечение двух бесконечных 2D-прямых.
        /// </summary>
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

        /// <summary>
        /// Псевдоскалярное произведение в 2D.
        /// </summary>
        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        /// <summary>
        /// Склеить коллинеарные линии в более длинные сегменты. Это уменьшает число лишних разбиений перед созданием стен.
        /// </summary>
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

            // Группируем по направлению, смещению и Z.
            var groups = data.GroupBy(x => new GenericAxisGroupKey(x.Dir, x.Offset, x.Z, tol)).ToList();
            List<Line> result = new List<Line>();

            foreach (var group in groups)
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

        /// <summary>
        /// Привести направление к каноническому виду, чтобы противоположные направления считались одинаковыми.
        /// </summary>
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

        /// <summary>
        /// Скалярное произведение в 2D.
        /// </summary>
        private static double Dot2D(XYZ a, XYZ b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        /// <summary>
        /// Построить линию из параметрического представления: направление + смещение + начало/конец интервала.
        /// </summary>
        private static Line BuildGenericAxisLine(XYZ dir, double offset, double z, double from, double to)
        {
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            XYZ p0 = new XYZ(dir.X * from + normal.X * offset, dir.Y * from + normal.Y * offset, z);
            XYZ p1 = new XYZ(dir.X * to + normal.X * offset, dir.Y * to + normal.Y * offset, z);

            return Line.CreateBound(p0, p1);
        }

        /// <summary>
        /// Удалить участки новых осей, которые уже перекрыты существующими стенами.
        /// </summary>
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

        /// <summary>
        /// Вычесть из одной новой линии все пересекающиеся с ней существующие стены.
        /// </summary>
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

                // Если линии лежат не на одной оси — пропускаем.
                double exOffset = Dot2D(ex.P0, normal);
                if (Math.Abs(exOffset - newOffset) > tol)
                    continue;

                // Переводим обе линии в параметрическую 1D-систему вдоль DIR.
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

            // Возвращаем оставшиеся после вычитания части как реальные линии Revit.
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

        /// <summary>
        /// Склеить пересекающиеся и соприкасающиеся интервалы.
        /// </summary>
        private static List<Interval1D> MergeIntervals(List<Interval1D> intervals)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (intervals == null || intervals.Count == 0)
                return result;

            List<Interval1D> ordered = intervals.Where(x => x != null && x.To - x.From > tol).OrderBy(x => x.From).ThenBy(x => x.To).ToList();

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

        /// <summary>
        /// Вычесть из исходного интервала список вырезаемых интервалов.
        /// </summary>
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

        /// <summary>
        /// Применить к стене параметры из пресета
        /// </summary>
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
    }
}
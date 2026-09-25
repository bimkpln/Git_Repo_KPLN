using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;          // ModelItem
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ClashApi = Autodesk.Navisworks.Api.Clash;
// НАШЛИ (по третьей DLL, Autodesk.Navisworks.ComApi.dll, которую вы
// добавили): настоящий публичный ComApiBridge всё-таки есть в 2020, просто
// лежит в отдельной сборке от Api.dll и Interop.ComApi.dll. State здесь уже
// строго типизирован как InwOpState10 (без каста), и главное — есть
// ComApiBridge.ToModelItem(InwOaPath), которым удобно и надёжно превращать
// путь клэша в обычный managed ModelItem вместо ручного обхода COM Nodes().
using Autodesk.Navisworks.Api.ComApi;

namespace KPLN_NavisMcpBridge
{
    internal sealed class ResourceNotFoundException : Exception
    {
        public ResourceNotFoundException(string message) : base(message) { }
    }

    internal sealed class StateConflictException : Exception
    {
        public StateConflictException(string message, Exception innerException = null)
            : base(message, innerException) { }
    }

    // === DTO для JSON ===

    public class ClashTestDto
    {
        public string Name;
        public string Status;       // nwEClashTestStatus: NEW/OLD/PARTIAL/OK
        public int ResultCount;
        // Ниже — диагностика для вопроса "почему у всех результатов
        // глубина пересечения больше N мм": если у теста настроен большой
        // Tolerance, мелкие пересечения в results() физически не попадают —
        // тест их с самого начала не репортит. Поля берутся из
        // InwOclClashTest2 (расширение базового InwOclClashTest), могут
        // быть null, если объект теста почему-то его не поддерживает.
        public double? Tolerance;      // допуск теста, единицы — как и Distance (см. ClashResultDto)
        public string ToleranceType;   // nwEClashTestToleranceType: Absolute/Relative и т.п.
        public string TestType;        // nwEClashTestType: Hard/Clearance/Duplicates и т.п.
    }

    public class Point3Dto { public double X; public double Y; public double Z; }

    public class BoundDto { public Point3Dto Min; public Point3Dto Max; }

    public class PlanGeometryDto
    {
        public double AxisX;
        public double AxisY;
        public double CenterX;
        public double CenterY;
        public double HalfLength;
        public double HalfThickness;
        public double Reliability;
        public int PointCount;
        public string Source;
        public double? NearestBroadFace;
        public double? NearestEndFace;
    }

    public class ClashResultDto
    {
        public string Name;
        public string Status;       // nwETestResultStatus: NEW/ACTIVE/APPROVED/RESOLVED/REVIEWED/UNMERGED
        public double Distance;
        public string ApprovedBy;
        public string Item1Path;
        public string Item2Path;
        public List<string> Comments;
        // Геометрия для группировки соседних коллизий и подсчёта габарита
        // объединённой зоны (напр. проёма под несколько труб) — единицы,
        // как и Distance (см. комментарий у GetResults), т.е. предположительно
        // метры (задокументированное поведение движка Navisworks: все
        // геометрические величины API — в метрах, независимо от единиц
        // файла/интерфейса). Pt1/Pt2 — точки контакта на первом/втором
        // элементе пары, Bound — bounding box ЗОНЫ ПЕРЕСЕЧЕНИЯ (маленький
        // локальный объём в месте контакта — НЕ путать с Item1Bound/
        // Item2Bound ниже).
        public Point3Dto Pt1;
        public Point3Dto Pt2;
        public BoundDto Bound;
        // Bounding box КАЖДОГО ЭЛЕМЕНТА ЦЕЛИКОМ (не зоны пересечения) —
        // ModelItem.Geometry.BoundingBox, подтверждено декомпиляцией
        // Autodesk.Navisworks.Api.dll (ModelItem.Geometry -> ModelGeometry,
        // ModelGeometry.BoundingBox -> BoundingBox3D с Min/Max Point3D).
        // Нужен, чтобы приблизительно оценить направление оси трубы: для
        // прямого сегмента трубы/изоляции самая длинная грань его
        // собственного bounding box'а обычно совпадает с направлением
        // прокладки (при условии, что сегмент проложен вдоль одной из
        // глобальных осей X/Y/Z — типичный случай для ортогональной
        // трассировки ИОС, но НЕ работает для диагональных участков).
        // Заполняется только если includeItemBounds=true — требует
        // резолва ModelItem, как и Item1Path/Item2Path.
        public BoundDto Item1Bound;
        public BoundDto Item2Bound;
        // Диагностика: откуда взят Item1Bound/Item2Bound — "leaf" (у самого
        // элемента клэша есть геометрия напрямую, HasGeometry=true) или
        // "union-of-N" (элемент клэша — составной узел БЕЗ собственной
        // геометрии, например тип/группа Revit; bbox посчитан как объединение
        // всех geometry-содержащих потомков — DescendantsAndSelf.Where(HasGeometry)).
        // "union-of-N" при большом N может означать, что путь клэша указывает
        // на объединяющий узел, а не на конкретный сегмент трубы — тогда
        // получившийся bbox может быть намного больше одного сегмента и не
        // годиться для оценки направления оси.
        public string Item1BoundSource;
        public string Item2BoundSource;
        public PlanGeometryDto Item1PlanGeometry;
        public PlanGeometryDto Item2PlanGeometry;

        /// <summary>
        /// Имя ГРУППЫ коллизий, в которую пользователь объединил этот
        /// результат в Clash Detective (вида "Группа #28"), либо null,
        /// если результат не сгруппирован. Берётся из InwOclTestResult2.
        /// GroupPath — интерфейс "2" есть у результатов Navisworks 2020,
        /// но старый InwOclTestResult этого свойства не имеет, поэтому
        /// делаем безопасный каст.
        ///
        /// Зачем: пучок инженерных элементов, идущих через ОДНО общее
        /// отверстие, определяется в первую очередь тем, что пользователь
        /// сам собрал их в одну группу. Расстоянием это не заменяется:
        /// проверено на отчёте в 96 результатов — кластеризация только по
        /// близости даёт 41 ложную коллизию, потому что склеивает
        /// воздуховоды, идущие через РАЗНЫЕ отверстия в одной стене.
        /// </summary>
        public string Group;

        /// <summary>
        /// Размер элемента, как он записан в свойствах ("ø200 мм-ø200 мм",
        /// "300x200-300x200"), а НЕ померенный по геометрии.
        ///
        /// Зачем: габарит bounding box'а систематически врёт про сечение.
        /// У круглого противопожарного клапана ø200 коробка даёт 253 мм
        /// (корпус), у отвода R=1.5 — вдвое больше диаметра, а у любого
        /// диагонального в плане элемента раздувается по обеим
        /// горизонтальным осям. Решение "нужно ли отверстие" принимается
        /// именно по номинальному сечению, поэтому оно нужно как есть.
        /// </summary>
        public string Item1Size;
        public string Item2Size;
        public string Item1Mark;
        public string Item2Mark;
        public string Item1Workset;
        public string Item2Workset;
    }

    public class ResultStatusBatchDto
    {
        public List<string> updated { get; set; } = new List<string>();
        public List<string> missing { get; set; } = new List<string>();
        public Dictionary<string, string> failed { get; set; } =
            new Dictionary<string, string>(StringComparer.Ordinal);
    }

    /// <summary>
    /// Обёртка над Autodesk.Navisworks.Api.Interop.ComApi (COM API) —
    /// единственный работающий путь к Clash Detective в Navisworks 2020.
    ///
    /// Всё ниже проверено декомпиляцией реальных Autodesk.Navisworks.Api.dll
    /// и Autodesk.Navisworks.Interop.ComApi.dll (метаданные ECMA-335,
    /// имена классов/методов/свойств и их реальные сигнатуры) — НЕ по памяти.
    /// Одно место, которое проверить статически было нельзя (нативный C++/CLI
    /// код без управляемых строковых констант) — это то, что тесты Clash
    /// Detective ищутся перебором плагинов текущего документа с приведением
    /// типа к InwOpClashElement, а не по жёстко заданному строковому ID —
    /// это осознанный выбор, устойчивый к любому реальному имени плагина.
    ///
    /// ВАЖНО про стиль API: часть "свойств" в этом COM-интерфейсе на самом
    /// деле обычные МЕТОДЫ без префикса get_ (Plugins(), Tests(), results(),
    /// Comments(), Nodes() — вызываются со скобками), а часть — настоящие
    /// C#-свойства (Path1, Path2, status, name, distance — без скобок).
    /// Перепутать легко, поэтому явно комментирую каждый вызов.
    /// </summary>
    internal static class ClashService
    {
        private static ComApi.InwOpState10 State => ComApiBridge.State;

        /// <summary>
        /// В Navisworks 2020 нет надёжного публичного способа получить
        /// элемент Clash Detective по строковому ID плагина (в отличие от
        /// более новых версий SDK). Поэтому просто перебираем все плагины
        /// текущего документа и берём первый, который приводится к
        /// InwOpClashElement — именно этот интерфейс отвечает за
        /// Clash Detective (Tests / RunAllTests / ClearResults).
        /// </summary>
        private static ComApi.InwOpClashElement GetClashElement()
        {
            foreach (object plugin in State.Plugins()) // Plugins() — метод, не свойство!
            {
                if (plugin is ComApi.InwOpClashElement clash)
                    return clash;
            }
            throw new StateConflictException(
                "Не найден элемент Clash Detective среди плагинов документа " +
                "(документ не открыт или Clash Detective недоступен)");
        }

        private static IEnumerable<object> Enumerate(System.Collections.IEnumerable coll)
        {
            foreach (var item in coll) yield return item;
        }

        public static List<ClashTestDto> ListTests()
        {
            var clash = GetClashElement();
            var list = new List<ClashTestDto>();
            foreach (var obj in Enumerate(clash.Tests())) // Tests() — метод!
            {
                if (obj is ComApi.InwOclClashTest test)
                {
                    var dto = new ClashTestDto
                    {
                        Name = test.name,               // свойство (get_name/set_name)
                        Status = test.status.ToString(),// свойство, enum nwEClashTestStatus
                        ResultCount = CountResults(test)
                    };
                    // Tolerance/ToleranceType/TestType — на InwOclClashTest2
                    // (расширение базового InwOclClashTest на том же COM-объекте).
                    if (test is ComApi.InwOclClashTest2 test2)
                    {
                        dto.Tolerance = test2.Tolerance;
                        dto.ToleranceType = test2.ToleranceType.ToString();
                        dto.TestType = test2.TestType.ToString();
                    }
                    list.Add(dto);
                }
            }
            return list;
        }

        private static int CountResults(ComApi.InwOclClashTest test)
        {
            int n = 0;
            foreach (var _ in Enumerate(test.results())) n++; // results() — метод!
            return n;
        }

        private static ComApi.InwOclClashTest FindTest(string name)
        {
            var clash = GetClashElement();
            foreach (var obj in Enumerate(clash.Tests()))
            {
                if (obj is ComApi.InwOclClashTest test && test.name == name)
                    return test;
            }
            throw new ResourceNotFoundException($"Тест Clash Detective '{name}' не найден");
        }

        private static ClashApi.DocumentClashTests FindManagedTest(
            string name,
            out ClashApi.ClashTest foundTest)
        {
            var document = Application.ActiveDocument;
            if (document == null)
                throw new StateConflictException("В Navisworks не открыт документ");

            var clash = ClashApi.DocumentClash.ClashInstance(document);
            if (clash == null || clash.TestsData == null)
                throw new StateConflictException("Clash Detective недоступен для текущего документа");

            foreach (SavedItem item in clash.TestsData.Tests)
            {
                var test = item as ClashApi.ClashTest;
                if (test != null && test.DisplayName == name)
                {
                    foundTest = test;
                    return clash.TestsData;
                }
            }

            throw new ResourceNotFoundException($"Тест Clash Detective '{name}' не найден");
        }

        private static void FindManagedGroups(
            GroupItem parent,
            HashSet<string> requested,
            Dictionary<string, List<ClashApi.ClashResultGroup>> found)
        {
            foreach (SavedItem item in parent.Children)
            {
                var group = item as ClashApi.ClashResultGroup;
                if (group == null) continue;

                if (requested.Contains(group.DisplayName))
                    found[group.DisplayName].Add(group);

                FindManagedGroups(group, requested, found);
            }
        }

        private static void FindManagedResults(
            GroupItem parent,
            HashSet<string> requested,
            Dictionary<string, List<ClashApi.ClashResult>> found)
        {
            foreach (SavedItem item in parent.Children)
            {
                var result = item as ClashApi.ClashResult;
                if (result != null && requested.Contains(result.DisplayName))
                    found[result.DisplayName].Add(result);

                var group = item as GroupItem;
                if (group != null)
                    FindManagedResults(group, requested, found);
            }
        }

        private static ClashApi.ClashResultStatus ParseManagedStatus(string status)
        {
            var normalized = status ?? string.Empty;
            const string comPrefix = "eTestResultStatus_";
            if (normalized.StartsWith(comPrefix, StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(comPrefix.Length);

            if (Enum.TryParse(normalized, true, out ClashApi.ClashResultStatus parsed))
                return parsed;

            throw new ArgumentException(
                $"Неизвестный статус: {status}. Допустимые: New, Active, Approved, Resolved, Reviewed");
        }

        public static List<ClashResultDto> GetResults(string testName, bool includePaths = true, string statusFilter = null, bool includeItemBounds = false, bool includeSizes = false, bool includeMarks = false, bool includeWorksets = false)
        {
            var test = FindTest(testName);
            var list = new List<ClashResultDto>();
            var planGeometryCache = new Dictionary<Guid, WallMeshAnalysis>();

            // Нормализуем фильтр так же, как в SetResultStatus: принимаем
            // и короткое имя ("Active"), и полное значение enum
            // ("eTestResultStatus_ACTIVE") — регистр не важен.
            string statusFilterFull = null;
            if (!string.IsNullOrEmpty(statusFilter))
            {
                statusFilterFull = statusFilter.StartsWith("eTestResultStatus_")
                    ? statusFilter.ToUpperInvariant()
                    : ("ETESTRESULTSTATUS_" + statusFilter.ToUpperInvariant());
            }

            foreach (var obj in Enumerate(test.results())) // results() — метод!
            {
                if (!(obj is ComApi.InwOclTestResult r)) continue;

                if (statusFilterFull != null &&
                    r.status.ToString().ToUpperInvariant() != statusFilterFull)
                    continue; // фильтр по статусу — пропускаем до построения DTO/геометрии

                var comments = new List<string>();
                foreach (var c in Enumerate(r.Comments())) // Comments() — метод!
                {
                    if (c is ComApi.InwOpComment comment)
                        comments.Add($"{comment.User}: {comment.Body}");
                }

                // Path1/Path2 — свойства. Резолвим ModelItem один раз и
                // переиспользуем и для пути (AncestorsAndSelf), и для
                // собственного bounding box'а элемента — так дороже не
                // становится, если нужно и то, и другое одновременно.
                Autodesk.Navisworks.Api.ModelItem item1 = null, item2 = null;
                if (includePaths || includeItemBounds || includeSizes || includeMarks || includeWorksets)
                {
                    item1 = ComApiBridge.ToModelItem(r.Path1);
                    item2 = ComApiBridge.ToModelItem(r.Path2);
                }

                var item1Path = includePaths ? DescribeModelItemPath(item1) : null;
                var item2Path = includePaths ? DescribeModelItemPath(item2) : null;
                var clashPoint = DescribePoint(r.Pt1);
                var dto = new ClashResultDto
                {
                    Name = r.name,
                    Status = r.status.ToString(),
                    Distance = r.distance,
                    ApprovedBy = r.ApprovedBy,
                    // DescribeModelItemPath ходит по AncestorsAndSelf — на
                    // тестах с тысячами результатов это дорого (десятки
                    // тысяч COM-вызовов на один запрос) и раньше приводило
                    // к таймауту моста. Пропускаем, если пути не нужны
                    // (includePaths=false) — например, при фильтрации
                    // только по дистанции/статусу.
                    Item1Path = item1Path,
                    Item2Path = item2Path,
                    Comments = comments,
                    Pt1 = clashPoint,
                    Pt2 = DescribePoint(r.Pt2),
                    Bound = DescribeBound(r.Bound),
                    Group = DescribeGroup(r)
                };
                list.Add(dto);

                if (includeItemBounds)
                {
                    var b1 = DescribeItemBoundingBox(item1);
                    var b2 = DescribeItemBoundingBox(item2);
                    list[list.Count - 1].Item1Bound = b1.Bound;
                    list[list.Count - 1].Item1BoundSource = b1.Source;
                    list[list.Count - 1].Item2Bound = b2.Bound;
                    list[list.Count - 1].Item2BoundSource = b2.Source;

                    if (IsWallPath(item1Path))
                        list[list.Count - 1].Item1PlanGeometry =
                            DescribeCachedPlanGeometry(item1, b1.Bound, clashPoint, planGeometryCache);
                    if (IsWallPath(item2Path))
                        list[list.Count - 1].Item2PlanGeometry =
                            DescribeCachedPlanGeometry(item2, b2.Bound, clashPoint, planGeometryCache);
                }

                if (includeSizes)
                {
                    list[list.Count - 1].Item1Size = DescribeModelItemGuiSize(item1);
                    list[list.Count - 1].Item2Size = DescribeModelItemGuiSize(item2);
                }

                if (includeMarks)
                {
                    list[list.Count - 1].Item1Mark = DescribeModelItemGuiMark(item1);
                    list[list.Count - 1].Item2Mark = DescribeModelItemGuiMark(item2);
                }

                if (includeWorksets)
                {
                    list[list.Count - 1].Item1Workset = DescribeModelItemGuiWorkset(item1);
                    list[list.Count - 1].Item2Workset = DescribeModelItemGuiWorkset(item2);
                }
            }
            return list;
        }

        /// <summary>
        /// Имя группы коллизий результата. У Navisworks 2020 оно доступно
        /// через расширенный интерфейс InwOclTestResult2.GroupPath (string).
        /// Каст безопасный: если реализация его не поддерживает, вернём null,
        /// а не уроним весь запрос.
        /// </summary>
        private static string DescribeGroup(ComApi.InwOclTestResult r)
        {
            try
            {
                var r2 = r as ComApi.InwOclTestResult2;
                if (r2 == null) return null;
                var g = r2.GroupPath;
                return string.IsNullOrWhiteSpace(g) ? null : g;
            }
            catch
            {
                return null;
            }
        }

        private static string DescribeTopLevelGroup(ComApi.InwOclTestResult r)
        {
            var path = DescribeGroup(r);
            if (string.IsNullOrWhiteSpace(path)) return null;

            var normalized = path.Replace("\r\n", "\n").Replace('\r', '\n');
            var separator = normalized.IndexOf('\n');
            if (separator <= 0) return null;

            var groupName = normalized.Substring(0, separator).Trim();
            return string.IsNullOrEmpty(groupName) ? null : groupName;
        }

        // Имена свойств, в которых лежит сечение. Ищем по ОТОБРАЖАЕМОМУ имени
        // и в русской, и в английской выгрузке. Категория "Объект" имеет
        // приоритет: именно там Navisworks показывает номинальный размер труб
        // и соединительных деталей (например, "Размер" = "ø80").
        private static readonly string[] ObjectCategoryNames = { "Объект", "Object" };
        private static readonly string[] SizeNames = { "Размер", "Общий размер", "Size", "Overall Size" };
        private static readonly string[] WidthNames = { "Ширина", "Width" };
        private static readonly string[] HeightNames = { "Высота", "Height" };
        private static readonly string[] DiameterNames = { "Диаметр", "Diameter" };
        private static readonly string[] MarkNames = { "Марка", "Mark" };
        private static readonly string[] WorksetNames =
        {
            "Рабочий набор", "Имя рабочего набора", "Workset", "Workset Name"
        };

        private static bool NameMatches(string name, string[] candidates)
        {
            if (string.IsNullOrWhiteSpace(name) || candidates == null) return false;

            name = name.Trim();
            foreach (var c in candidates)
            {
                if (string.IsNullOrWhiteSpace(c)) continue;
                if (string.Equals(name, c.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string PropertyText(Autodesk.Navisworks.Api.DataProperty p)
        {
            try
            {
                var v = p.Value;
                return v == null ? null : v.ToDisplayString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Читает размер из тех же GUI-свойств COM-пути, которые Navisworks
        /// показывает справа для выбранного clash-элемента: вкладка
        /// "Объект" -> строка "Размер" -> значение вроде "ø80".
        /// </summary>
        private static string DescribeModelItemGuiSize(ModelItem item)
        {
            if (item == null) return null;

            try
            {
                // Clash часто указывает на дочерний геометрический узел, а
                // вкладка "Объект" принадлежит его родительскому Revit-
                // элементу. Идём от самого узла вверх до первого размера.
                foreach (var candidate in item.AncestorsAndSelf)
                {
                    var size = DescribeSingleModelItemGuiSize(candidate);
                    if (!string.IsNullOrWhiteSpace(size)) return size;
                }
            }
            catch
            {
                // Отдельный битый путь не должен ронять весь отчёт.
            }

            return null;
        }

        private static string DescribeSingleModelItemGuiSize(ModelItem item)
        {
            // Нормализуем путь тем же способом, что и официальные примеры:
            // ModelItem -> ComApiBridge.ToInwOaPath -> GUI property node.
            var path = ComApiBridge.ToInwOaPath(item);
            var node = State.GetGUIPropertyNode(path, true);
            var attributes = node.GUIAttributes();

            foreach (var obj in Enumerate(attributes))
            {
                var attribute = obj as ComApi.InwGUIAttribute2;
                if (attribute == null ||
                    !NameMatches(attribute.ClassUserName, ObjectCategoryNames)) continue;

                var size = SizeFromGuiAttribute(attribute);
                if (!string.IsNullOrWhiteSpace(size)) return size;
            }

            // Fallback для моделей, где та же строка опубликована на
            // другой вкладке, но по-прежнему называется "Размер".
            foreach (var obj in Enumerate(attributes))
            {
                var attribute = obj as ComApi.InwGUIAttribute2;
                if (attribute == null ||
                    NameMatches(attribute.ClassUserName, ObjectCategoryNames)) continue;

                var size = SizeFromGuiAttribute(attribute);
                if (!string.IsNullOrWhiteSpace(size)) return size;
            }

            return null;
        }

        private static string SizeFromGuiAttribute(ComApi.InwGUIAttribute2 attribute)
        {
            foreach (var obj in Enumerate(attribute.Properties()))
            {
                var property = obj as ComApi.InwOaProperty;
                if (property == null) continue;

                var value = property.value == null ? null : property.value.ToString();
                if (string.IsNullOrWhiteSpace(value)) continue;

                var propertyName = (property.UserName ?? string.Empty).Trim();
                var isNamedSize = NameMatches(propertyName, SizeNames);
                var isDiameterSize =
                    propertyName.IndexOf("размер", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (value.IndexOf('ø') >= 0 || value.IndexOf('Ø') >= 0 || value.IndexOf('⌀') >= 0);
                if (isNamedSize || isDiameterSize) return value.Trim();
            }

            return null;
        }

        private static string DescribeModelItemGuiMark(ModelItem item)
        {
            if (item == null) return null;

            try
            {
                foreach (var candidate in item.AncestorsAndSelf)
                {
                    var value = DescribeSingleModelItemGuiProperty(candidate, MarkNames);
                    if (!string.IsNullOrWhiteSpace(value)) return value;

                    // У части Revit-элементов стандартная "Марка" доступна
                    // через managed PropertyCategories, хотя отсутствует в
                    // GUIPropertyNode того же COM-пути.
                    value = DescribeSingleModelItemProperty(candidate, MarkNames);
                    if (!string.IsNullOrWhiteSpace(value)) return value;
                }
            }
            catch
            {
                // Отдельный битый путь не должен ронять весь отчёт.
            }

            return null;
        }

        private static string DescribeModelItemGuiWorkset(ModelItem item)
        {
            if (item == null) return null;

            try
            {
                // Clash может ссылаться на геометрию ниже Revit-элемента.
                // Рабочий набор при этом показывается на вкладке "Объект"
                // одного из родительских узлов, поэтому поднимаемся до корня.
                foreach (var candidate in item.AncestorsAndSelf)
                {
                    var value = DescribeSingleModelItemGuiProperty(candidate, WorksetNames);
                    if (!string.IsNullOrWhiteSpace(value)) return value;

                    value = DescribeSingleModelItemProperty(candidate, WorksetNames);
                    if (!string.IsNullOrWhiteSpace(value)) return value;
                }
            }
            catch
            {
                // Отдельный битый путь не должен ронять весь отчёт.
            }

            return null;
        }

        private static string DescribeSingleModelItemProperty(
            ModelItem item,
            string[] propertyNames)
        {
            foreach (var category in item.PropertyCategories)
            {
                foreach (var property in category.Properties)
                {
                    if (!NameMatches(property.DisplayName, propertyNames)) continue;

                    var value = PropertyText(property);
                    if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
                }
            }

            return null;
        }

        private static string DescribeSingleModelItemGuiProperty(
            ModelItem item,
            string[] propertyNames)
        {
            var path = ComApiBridge.ToInwOaPath(item);
            var node = State.GetGUIPropertyNode(path, true);
            var attributes = node.GUIAttributes();

            foreach (var obj in Enumerate(attributes))
            {
                var attribute = obj as ComApi.InwGUIAttribute2;
                if (attribute == null ||
                    !NameMatches(attribute.ClassUserName, ObjectCategoryNames)) continue;

                var value = ValueFromGuiAttribute(attribute, propertyNames);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }

            foreach (var obj in Enumerate(attributes))
            {
                var attribute = obj as ComApi.InwGUIAttribute2;
                if (attribute == null ||
                    NameMatches(attribute.ClassUserName, ObjectCategoryNames)) continue;

                var value = ValueFromGuiAttribute(attribute, propertyNames);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }

            return null;
        }

        private static string ValueFromGuiAttribute(
            ComApi.InwGUIAttribute2 attribute,
            string[] propertyNames)
        {
            foreach (var obj in Enumerate(attribute.Properties()))
            {
                var property = obj as ComApi.InwOaProperty;
                if (property == null || !NameMatches(property.UserName, propertyNames)) continue;

                var value = property.value == null ? null : property.value.ToString();
                if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
            }

            return null;
        }

        /// <summary>
        /// Размер элемента из его свойств. Сначала ищем готовую строку
        /// "Размер" ("ø200 мм-ø200 мм", "300x200-300x200"); если её нет —
        /// собираем из Диаметра либо Ширины и Высоты. Если у самого узла
        /// свойств нет (он контейнер), проверяем всех потомков: вкладка
        /// "Объект" может находиться на промежуточном узле без геометрии.
        /// </summary>
        private static string DescribeItemSize(Autodesk.Navisworks.Api.ModelItem item)
        {
            if (item == null) return null;
            var direct = SizeFromProperties(item);
            if (!string.IsNullOrEmpty(direct)) return direct;

            try
            {
                foreach (var child in item.Descendants)
                {
                    var s2 = SizeFromProperties(child);
                    if (!string.IsNullOrEmpty(s2)) return s2;
                }
            }
            catch { /* ignore */ }

            return null;
        }

        private static string SizeFromProperties(Autodesk.Navisworks.Api.ModelItem item)
        {
            string width = null, height = null, diameter = null;
            try
            {
                // Сначала строго проверяем вкладку "Объект": в ней лежит
                // значение, которое видит пользователь в Navisworks.
                foreach (var cat in item.PropertyCategories)
                {
                    if (!NameMatches(cat.DisplayName, ObjectCategoryNames)) continue;
                    var objectSize = SizeFromCategory(cat, ref width, ref height, ref diameter);
                    if (!string.IsNullOrWhiteSpace(objectSize)) return objectSize;
                }

                // Fallback для семейств, которые публикуют размер на другой
                // вкладке (например, "Элемент" или "Тип").
                foreach (var cat in item.PropertyCategories)
                {
                    if (NameMatches(cat.DisplayName, ObjectCategoryNames)) continue;
                    var size = SizeFromCategory(cat, ref width, ref height, ref diameter);
                    if (!string.IsNullOrWhiteSpace(size)) return size;
                }
            }
            catch { /* свойства бывают недоступны — не роняем весь запрос */ }

            if (!string.IsNullOrWhiteSpace(diameter)) return "ø" + diameter;
            if (!string.IsNullOrWhiteSpace(width) && !string.IsNullOrWhiteSpace(height))
                return width + "x" + height;
            return null;
        }

        private static string SizeFromCategory(
            Autodesk.Navisworks.Api.PropertyCategory category,
            ref string width,
            ref string height,
            ref string diameter)
        {
            foreach (var p in category.Properties)
            {
                var name = p.DisplayName;
                if (string.IsNullOrWhiteSpace(name)) continue;

                if (NameMatches(name, SizeNames))
                {
                    var value = PropertyText(p);
                    if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
                }
                else if (diameter == null && NameMatches(name, DiameterNames))
                    diameter = PropertyText(p);
                else if (width == null && NameMatches(name, WidthNames))
                    width = PropertyText(p);
                else if (height == null && NameMatches(name, HeightNames))
                    height = PropertyText(p);
            }

            return null;
        }

        private static string DescribeModelItemPath(Autodesk.Navisworks.Api.ModelItem item)
        {
            if (item == null) return null;
            return string.Join(" / ", item.AncestorsAndSelf.Select(a => a.DisplayName));
        }

        private static bool IsWallPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            return path.IndexOf("/ Стены /", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/ Несущие стены /", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("Базовая стена", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private sealed class Triangle3
        {
            public Point3Dto A;
            public Point3Dto B;
            public Point3Dto C;
        }

        private sealed class WallMeshAnalysis
        {
            public PlanGeometryDto Plan;
            public List<Triangle3> Triangles;
        }

        private sealed class PrimitivePointCollector : ComApi.InwSimplePrimitivesCB
        {
            private readonly double[] _matrix;
            private readonly int _maxPoints;

            public readonly List<Point3Dto> ColumnVectorPoints = new List<Point3Dto>();
            public readonly List<Point3Dto> RowVectorPoints = new List<Point3Dto>();
            public readonly List<Triangle3> ColumnVectorTriangles = new List<Triangle3>();
            public readonly List<Triangle3> RowVectorTriangles = new List<Triangle3>();

            public PrimitivePointCollector(double[] matrix, int maxPoints)
            {
                _matrix = matrix;
                _maxPoints = maxPoints;
            }

            public void Triangle(
                ComApi.InwSimpleVertex v1,
                ComApi.InwSimpleVertex v2,
                ComApi.InwSimpleVertex v3)
            {
                if (ColumnVectorPoints.Count + 3 > _maxPoints) return;
                var a = ReadPoint(v1);
                var b = ReadPoint(v2);
                var c = ReadPoint(v3);
                if (a == null || b == null || c == null) return;

                var ca = TransformColumnVector(a.X, a.Y, a.Z, _matrix);
                var cb = TransformColumnVector(b.X, b.Y, b.Z, _matrix);
                var cc = TransformColumnVector(c.X, c.Y, c.Z, _matrix);
                var ra = TransformRowVector(a.X, a.Y, a.Z, _matrix);
                var rb = TransformRowVector(b.X, b.Y, b.Z, _matrix);
                var rc = TransformRowVector(c.X, c.Y, c.Z, _matrix);

                ColumnVectorPoints.Add(ca);
                ColumnVectorPoints.Add(cb);
                ColumnVectorPoints.Add(cc);
                RowVectorPoints.Add(ra);
                RowVectorPoints.Add(rb);
                RowVectorPoints.Add(rc);
                ColumnVectorTriangles.Add(new Triangle3 { A = ca, B = cb, C = cc });
                RowVectorTriangles.Add(new Triangle3 { A = ra, B = rb, C = rc });
            }

            public void Line(ComApi.InwSimpleVertex v1, ComApi.InwSimpleVertex v2) { }
            public void Point(ComApi.InwSimpleVertex v1) { }
            public void SnapPoint(ComApi.InwSimpleVertex v1) { }

            private static Point3Dto ReadPoint(ComApi.InwSimpleVertex vertex)
            {
                var coordinates = vertex.coord as Array;
                if (coordinates == null || coordinates.Length < 3) return null;
                var first = coordinates.GetLowerBound(0);
                return new Point3Dto
                {
                    X = Convert.ToDouble(coordinates.GetValue(first)),
                    Y = Convert.ToDouble(coordinates.GetValue(first + 1)),
                    Z = Convert.ToDouble(coordinates.GetValue(first + 2))
                };
            }

            private static Point3Dto TransformColumnVector(double x, double y, double z, double[] m)
            {
                if (m == null || m.Length < 16)
                    return new Point3Dto { X = x, Y = y, Z = z };
                var w = m[12] * x + m[13] * y + m[14] * z + m[15];
                if (Math.Abs(w) < 1e-12) w = 1.0;
                return new Point3Dto
                {
                    X = (m[0] * x + m[1] * y + m[2] * z + m[3]) / w,
                    Y = (m[4] * x + m[5] * y + m[6] * z + m[7]) / w,
                    Z = (m[8] * x + m[9] * y + m[10] * z + m[11]) / w
                };
            }

            private static Point3Dto TransformRowVector(double x, double y, double z, double[] m)
            {
                if (m == null || m.Length < 16)
                    return new Point3Dto { X = x, Y = y, Z = z };
                var w = x * m[3] + y * m[7] + z * m[11] + m[15];
                if (Math.Abs(w) < 1e-12) w = 1.0;
                return new Point3Dto
                {
                    X = (x * m[0] + y * m[4] + z * m[8] + m[12]) / w,
                    Y = (x * m[1] + y * m[5] + z * m[9] + m[13]) / w,
                    Z = (x * m[2] + y * m[6] + z * m[10] + m[14]) / w
                };
            }
        }

        private static PlanGeometryDto DescribeCachedPlanGeometry(
            ModelItem item,
            BoundDto itemBound,
            Point3Dto clashPoint,
            Dictionary<Guid, WallMeshAnalysis> cache)
        {
            if (item == null || itemBound == null) return null;

            Guid key;
            try { key = item.InstanceGuid; }
            catch { key = Guid.Empty; }

            WallMeshAnalysis analysis = null;
            if (key != Guid.Empty) cache.TryGetValue(key, out analysis);
            if (analysis == null)
            {
                analysis = DescribePlanGeometry(item, itemBound);
                if (key != Guid.Empty) cache[key] = analysis;
            }
            return DescribeLocalWallFaces(analysis, clashPoint);
        }

        private static WallMeshAnalysis DescribePlanGeometry(ModelItem item, BoundDto itemBound)
        {
            const int maxPoints = 30000;
            var columnPoints = new List<Point3Dto>();
            var rowPoints = new List<Point3Dto>();
            var columnTriangles = new List<Triangle3>();
            var rowTriangles = new List<Triangle3>();

            try
            {
                foreach (var leaf in item.DescendantsAndSelf)
                {
                    if (!leaf.HasGeometry || columnPoints.Count >= maxPoints) continue;
                    var path = ComApiBridge.ToInwOaPath(leaf);
                    if (path == null) continue;

                    foreach (var obj in Enumerate(path.Fragments()))
                    {
                        var fragment = obj as ComApi.InwOaFragment3;
                        if (fragment == null || columnPoints.Count >= maxPoints) continue;

                        var matrix = ReadMatrix(fragment.GetLocalToWorldMatrix());
                        var collector = new PrimitivePointCollector(
                            matrix, maxPoints - columnPoints.Count);
                        fragment.GenerateSimplePrimitives(
                            ComApi.nwEVertexProperty.eNONE, collector);
                        columnPoints.AddRange(collector.ColumnVectorPoints);
                        rowPoints.AddRange(collector.RowVectorPoints);
                        columnTriangles.AddRange(collector.ColumnVectorTriangles);
                        rowTriangles.AddRange(collector.RowVectorTriangles);
                    }
                }
            }
            catch
            {
                return null;
            }

            if (columnPoints.Count < 6) return null;
            var useColumn = BoundingScore(columnPoints, itemBound) <= BoundingScore(rowPoints, itemBound);
            return new WallMeshAnalysis
            {
                Plan = FitPlanGeometry(useColumn ? columnPoints : rowPoints,
                    useColumn ? "mesh-pca-column" : "mesh-pca-row"),
                Triangles = useColumn ? columnTriangles : rowTriangles
            };
        }

        private static PlanGeometryDto DescribeLocalWallFaces(
            WallMeshAnalysis analysis,
            Point3Dto clashPoint)
        {
            if (analysis == null || analysis.Plan == null) return null;
            var source = analysis.Plan;
            var result = new PlanGeometryDto
            {
                AxisX = source.AxisX,
                AxisY = source.AxisY,
                CenterX = source.CenterX,
                CenterY = source.CenterY,
                HalfLength = source.HalfLength,
                HalfThickness = source.HalfThickness,
                Reliability = source.Reliability,
                PointCount = source.PointCount,
                Source = source.Source
            };
            if (clashPoint == null || analysis.Triangles == null) return result;

            var wallNormalX = -source.AxisY;
            var wallNormalY = source.AxisX;
            var nearestBroad = double.PositiveInfinity;
            var nearestEnd = double.PositiveInfinity;

            foreach (var triangle in analysis.Triangles)
            {
                var ab = Subtract(triangle.B, triangle.A);
                var ac = Subtract(triangle.C, triangle.A);
                var normal = Cross(ab, ac);
                var normalLength = Math.Sqrt(Dot(normal, normal));
                if (normalLength <= 1e-12) continue;

                var horizontal = Math.Sqrt(normal.X * normal.X + normal.Y * normal.Y);
                if (horizontal / normalLength < 0.7) continue;
                var nx = normal.X / horizontal;
                var ny = normal.Y / horizontal;
                var alongWall = Math.Abs(nx * source.AxisX + ny * source.AxisY);
                var acrossWall = Math.Abs(nx * wallNormalX + ny * wallNormalY);
                var distance = Math.Sqrt(PointTriangleDistanceSquared(
                    clashPoint, triangle.A, triangle.B, triangle.C));

                if (alongWall > acrossWall)
                    nearestEnd = Math.Min(nearestEnd, distance);
                else
                    nearestBroad = Math.Min(nearestBroad, distance);
            }

            if (!double.IsInfinity(nearestBroad)) result.NearestBroadFace = nearestBroad;
            if (!double.IsInfinity(nearestEnd)) result.NearestEndFace = nearestEnd;
            return result;
        }

        private static Point3Dto Subtract(Point3Dto a, Point3Dto b)
        {
            return new Point3Dto { X = a.X - b.X, Y = a.Y - b.Y, Z = a.Z - b.Z };
        }

        private static Point3Dto Add(Point3Dto a, Point3Dto b, double scale)
        {
            return new Point3Dto
            {
                X = a.X + b.X * scale,
                Y = a.Y + b.Y * scale,
                Z = a.Z + b.Z * scale
            };
        }

        private static double Dot(Point3Dto a, Point3Dto b)
        {
            return a.X * b.X + a.Y * b.Y + a.Z * b.Z;
        }

        private static Point3Dto Cross(Point3Dto a, Point3Dto b)
        {
            return new Point3Dto
            {
                X = a.Y * b.Z - a.Z * b.Y,
                Y = a.Z * b.X - a.X * b.Z,
                Z = a.X * b.Y - a.Y * b.X
            };
        }

        private static double DistanceSquared(Point3Dto a, Point3Dto b)
        {
            var d = Subtract(a, b);
            return Dot(d, d);
        }

        private static double PointTriangleDistanceSquared(
            Point3Dto p,
            Point3Dto a,
            Point3Dto b,
            Point3Dto c)
        {
            var ab = Subtract(b, a);
            var ac = Subtract(c, a);
            var ap = Subtract(p, a);
            var d1 = Dot(ab, ap);
            var d2 = Dot(ac, ap);
            if (d1 <= 0.0 && d2 <= 0.0) return DistanceSquared(p, a);

            var bp = Subtract(p, b);
            var d3 = Dot(ab, bp);
            var d4 = Dot(ac, bp);
            if (d3 >= 0.0 && d4 <= d3) return DistanceSquared(p, b);

            var vc = d1 * d4 - d3 * d2;
            if (vc <= 0.0 && d1 >= 0.0 && d3 <= 0.0)
            {
                var v = d1 / (d1 - d3);
                return DistanceSquared(p, Add(a, ab, v));
            }

            var cp = Subtract(p, c);
            var d5 = Dot(ab, cp);
            var d6 = Dot(ac, cp);
            if (d6 >= 0.0 && d5 <= d6) return DistanceSquared(p, c);

            var vb = d5 * d2 - d1 * d6;
            if (vb <= 0.0 && d2 >= 0.0 && d6 <= 0.0)
            {
                var w = d2 / (d2 - d6);
                return DistanceSquared(p, Add(a, ac, w));
            }

            var va = d3 * d6 - d5 * d4;
            if (va <= 0.0 && (d4 - d3) >= 0.0 && (d5 - d6) >= 0.0)
            {
                var edge = Subtract(c, b);
                var w = (d4 - d3) / ((d4 - d3) + (d5 - d6));
                return DistanceSquared(p, Add(b, edge, w));
            }

            var denominator = 1.0 / (va + vb + vc);
            var insideV = vb * denominator;
            var insideW = vc * denominator;
            var closest = Add(Add(a, ab, insideV), ac, insideW);
            return DistanceSquared(p, closest);
        }

        private static double[] ReadMatrix(ComApi.InwLTransform3f transform)
        {
            try
            {
                var transform3 = (ComApi.InwLTransform3f3)(object)transform;
                var values = transform3.Matrix as Array;
                if (values == null || values.Length < 16) return null;

                var first = values.GetLowerBound(0);
                var matrix = new double[16];
                for (var i = 0; i < matrix.Length; i++)
                    matrix[i] = Convert.ToDouble(values.GetValue(first + i));
                return matrix;
            }
            catch
            {
                return null;
            }
        }

        private static double BoundingScore(IList<Point3Dto> points, BoundDto target)
        {
            if (points == null || points.Count == 0 || target == null) return double.PositiveInfinity;
            var minX = points.Min(p => p.X);
            var minY = points.Min(p => p.Y);
            var minZ = points.Min(p => p.Z);
            var maxX = points.Max(p => p.X);
            var maxY = points.Max(p => p.Y);
            var maxZ = points.Max(p => p.Z);
            return Math.Abs(minX - target.Min.X) + Math.Abs(minY - target.Min.Y) +
                   Math.Abs(minZ - target.Min.Z) + Math.Abs(maxX - target.Max.X) +
                   Math.Abs(maxY - target.Max.Y) + Math.Abs(maxZ - target.Max.Z);
        }

        private static PlanGeometryDto FitPlanGeometry(IList<Point3Dto> points, string source)
        {
            if (points == null || points.Count < 6) return null;

            var meanX = points.Average(p => p.X);
            var meanY = points.Average(p => p.Y);
            double xx = 0.0, xy = 0.0, yy = 0.0;
            foreach (var point in points)
            {
                var dx = point.X - meanX;
                var dy = point.Y - meanY;
                xx += dx * dx;
                xy += dx * dy;
                yy += dy * dy;
            }
            xx /= points.Count;
            xy /= points.Count;
            yy /= points.Count;

            var trace = xx + yy;
            var root = Math.Sqrt(Math.Max(0.0, (xx - yy) * (xx - yy) + 4.0 * xy * xy));
            var major = (trace + root) * 0.5;
            var minor = (trace - root) * 0.5;
            if (major <= 1e-12) return null;

            var angle = 0.5 * Math.Atan2(2.0 * xy, xx - yy);
            var axisX = Math.Cos(angle);
            var axisY = Math.Sin(angle);
            var normalX = -axisY;
            var normalY = axisX;

            var minAlong = double.PositiveInfinity;
            var maxAlong = double.NegativeInfinity;
            var minAcross = double.PositiveInfinity;
            var maxAcross = double.NegativeInfinity;
            foreach (var point in points)
            {
                var along = point.X * axisX + point.Y * axisY;
                var across = point.X * normalX + point.Y * normalY;
                if (along < minAlong) minAlong = along;
                if (along > maxAlong) maxAlong = along;
                if (across < minAcross) minAcross = across;
                if (across > maxAcross) maxAcross = across;
            }

            var centerAlong = (minAlong + maxAlong) * 0.5;
            var centerAcross = (minAcross + maxAcross) * 0.5;
            return new PlanGeometryDto
            {
                AxisX = axisX,
                AxisY = axisY,
                CenterX = centerAlong * axisX + centerAcross * normalX,
                CenterY = centerAlong * axisY + centerAcross * normalY,
                HalfLength = (maxAlong - minAlong) * 0.5,
                HalfThickness = (maxAcross - minAcross) * 0.5,
                Reliability = major / Math.Max(minor, 1e-12),
                PointCount = points.Count,
                Source = source
            };
        }

        private sealed class ItemBoundResult
        {
            public BoundDto Bound;
            public string Source; // "leaf" | "union-of-N" | null (нет геометрии вовсе)
        }

        // ModelItem.Geometry.BoundingBox — подтверждено декомпиляцией
        // Autodesk.Navisworks.Api.dll (ModelGeometry.BoundingBox ->
        // BoundingBox3D, поля Min/Max -> Point3D X/Y/Z, те же типы double,
        // что и у InwLBox3f из COM API).
        //
        // ВАЖНО (обнаружено на реальных данных): элемент, на который
        // указывает путь клэша (Path1/Path2 -> ToModelItem), сам по себе
        // ЧАСТО не несёт геометрию напрямую (HasGeometry=false) — это
        // составной узел (тип/группа, IsComposite), а фактическая геометрия
        // лежит на его потомках. В этом случае берём объединение bounding
        // box'ов всех потомков с HasGeometry=true (DescendantsAndSelf) —
        // это единственный практический способ получить хоть какой-то
        // bbox для составного узла без разбора примитивов геометрии
        // (триангуляции) вручную. Source="union-of-N" в результате — сигнал
        // вызывающей стороне, что bbox может быть шире одного сегмента
        // трубы, если под узлом объединено несколько разных элементов.
        private static ItemBoundResult DescribeItemBoundingBox(Autodesk.Navisworks.Api.ModelItem item)
        {
            if (item == null) return new ItemBoundResult { Bound = null, Source = null };

            double minX = double.PositiveInfinity, minY = double.PositiveInfinity, minZ = double.PositiveInfinity;
            double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity, maxZ = double.NegativeInfinity;
            int count = 0;

            if (item.HasGeometry)
            {
                var bb = item.Geometry.BoundingBox;
                if (!bb.IsEmpty)
                {
                    minX = bb.Min.X; minY = bb.Min.Y; minZ = bb.Min.Z;
                    maxX = bb.Max.X; maxY = bb.Max.Y; maxZ = bb.Max.Z;
                    count = 1;
                }
            }
            else
            {
                foreach (var leaf in item.DescendantsAndSelf)
                {
                    if (!leaf.HasGeometry) continue;
                    var bb = leaf.Geometry.BoundingBox;
                    if (bb.IsEmpty) continue;
                    count++;
                    if (bb.Min.X < minX) minX = bb.Min.X;
                    if (bb.Min.Y < minY) minY = bb.Min.Y;
                    if (bb.Min.Z < minZ) minZ = bb.Min.Z;
                    if (bb.Max.X > maxX) maxX = bb.Max.X;
                    if (bb.Max.Y > maxY) maxY = bb.Max.Y;
                    if (bb.Max.Z > maxZ) maxZ = bb.Max.Z;
                }
            }

            if (count == 0)
            {
                // Диагностика: и сам узел без геометрии, и среди
                // DescendantsAndSelf ни одного geometry-узла не нашлось —
                // значит либо это реально пустой узел (маловероятно для
                // элемента, участвовавшего в клэше), либо геометрия у него
                // организована не через дерево потомков, а через отдельный
                // механизм инстансирования (Instances/IsInsert) — дампим
                // сырые факты, чтобы понять, что это за узел, не гадая.
                int childCount = 0, descCount = 0, instCount = -1;
                try { childCount = item.Children.Count(); } catch { childCount = -1; }
                try { descCount = item.Descendants.Count(); } catch { descCount = -1; }
                try { instCount = item.Instances.Count(); } catch { instCount = -1; }
                var diag = $"empty(HasGeometry={item.HasGeometry},IsComposite={item.IsComposite}," +
                           $"IsInsert={item.IsInsert},IsCollection={item.IsCollection}," +
                           $"Children={childCount},Descendants={descCount},Instances={instCount}," +
                           $"DisplayName={item.DisplayName})";
                return new ItemBoundResult { Bound = null, Source = diag };
            }

            return new ItemBoundResult
            {
                Bound = new BoundDto
                {
                    Min = new Point3Dto { X = minX, Y = minY, Z = minZ },
                    Max = new Point3Dto { X = maxX, Y = maxY, Z = maxZ }
                },
                Source = item.HasGeometry ? "leaf" : $"union-of-{count}"
            };
        }

        private static Point3Dto DescribePoint(ComApi.InwLPos3f p)
        {
            if (p == null) return null;
            // data1/data2/data3 — X/Y/Z (double), подтверждено декомпиляцией
            // Autodesk.Navisworks.Interop.ComApi.dll (тип InwLPos3f).
            return new Point3Dto { X = p.data1, Y = p.data2, Z = p.data3 };
        }

        private static BoundDto DescribeBound(ComApi.InwLBox3f b)
        {
            if (b == null) return null;
            // min_pos/max_pos — подтверждено декомпиляцией (тип InwLBox3f).
            return new BoundDto { Min = DescribePoint(b.min_pos), Max = DescribePoint(b.max_pos) };
        }

        public static void RunTest(string testName)
        {
            var test = FindTest(testName);
            // RunTest(int test_ndx, object ovProgress) — сигнатура подтверждена
            // декомпиляцией, но смысл test_ndx на объекте-тесте (а не коллекции)
            // непонятен из одних метаданных; похоже на неиспользуемый параметр
            // общего шаблона COM-интерфейса. ovProgress — необязательный колбэк
            // прогресса, null. Собрано и проверено: реальная сигнатура была
            // (int, object), а не () — это выяснилось только при попытке сборки.
            test.RunTest(0, null);
        }

        public static void SetResultStatus(string testName, string resultName, string status, string comment)
        {
            var test = FindTest(testName);
            ComApi.InwOclTestResult found = null;
            foreach (var obj in Enumerate(test.results()))
            {
                if (obj is ComApi.InwOclTestResult r && r.name == resultName)
                {
                    found = r;
                    break;
                }
            }
            if (found == null)
                throw new ResourceNotFoundException($"Результат '{resultName}' в тесте '{testName}' не найден");

            if (!string.IsNullOrEmpty(status))
            {
                // Значения enum COM-имперфейса сохраняют префикс eTestResultStatus_,
                // например "eTestResultStatus_APPROVED" — подтверждено декомпиляцией.
                var full = status.StartsWith("eTestResultStatus_") ? status : "eTestResultStatus_" + status;
                if (Enum.TryParse<ComApi.nwETestResultStatus>(full, true, out var parsed))
                    found.status = parsed;
                else
                    throw new ArgumentException(
                        $"Неизвестный статус: {status}. Допустимые: New, Active, Approved, Resolved, Reviewed");
            }

            if (!string.IsNullOrEmpty(comment))
            {
                // ВНИМАНИЕ — не проверено запуском: создание нового InwOpComment
                // в 2020 SDK, скорее всего, идёт через State.ObjectFactory(...),
                // а не напрямую "new". Это единственная операция в файле, которую
                // я не смог до конца проверить статически (нужен реальный тест
                // внутри Navisworks) — сигнатуру ObjectFactory видно в металданных,
                // но её enum-параметры (какой nwEObjectType выбрать) — нет.
                throw new NotImplementedException(
                    "Добавление комментария к результату клэша: нужно опробовать " +
                    "State.ObjectFactory(...) внутри реального Navisworks 2020 — " +
                    "статически проверить нельзя (см. README)");
            }
        }

        public static ResultStatusBatchDto SetResultStatuses(
            string testName,
            IDictionary<string, string> updates)
        {
            if (updates == null || updates.Count == 0)
                throw new ArgumentException("Не передано ни одного изменения статуса");

            var output = new ResultStatusBatchDto();
            var parsed = new Dictionary<string, ClashApi.ClashResultStatus>(StringComparer.Ordinal);
            foreach (var pair in updates)
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || string.IsNullOrWhiteSpace(pair.Value))
                {
                    output.failed[pair.Key ?? string.Empty] = "Отсутствует resultName или status";
                    continue;
                }

                try
                {
                    parsed[pair.Key.Trim()] = ParseManagedStatus(pair.Value);
                }
                catch (ArgumentException ex)
                {
                    output.failed[pair.Key] = ex.Message;
                }
            }

            if (parsed.Count == 0) return output;

            var testsData = FindManagedTest(testName, out var test);
            var requested = new HashSet<string>(parsed.Keys, StringComparer.Ordinal);
            var found = requested.ToDictionary(
                name => name,
                name => new List<ClashApi.ClashResult>(),
                StringComparer.Ordinal);
            FindManagedResults(test, requested, found);

            var pending = new List<KeyValuePair<ClashApi.ClashResult, ClashApi.ClashResultStatus>>();
            foreach (var pair in found)
            {
                if (pair.Value.Count == 0)
                {
                    output.missing.Add(pair.Key);
                    continue;
                }

                if (pair.Value.Count > 1)
                {
                    output.failed[pair.Key] =
                        $"В тесте найдено несколько результатов с именем '{pair.Key}'";
                    continue;
                }

                var result = pair.Value[0];
                var status = parsed[pair.Key];
                if (result.Status == status)
                {
                    output.updated.Add(pair.Key);
                    continue;
                }

                pending.Add(new KeyValuePair<ClashApi.ClashResult, ClashApi.ClashResultStatus>(
                    result, status));
            }

            if (pending.Count == 0) return output;

            var document = Application.ActiveDocument;
            using (var transaction = document.BeginTransaction("KPLN: пакетная смена статусов коллизий"))
            {
                foreach (var pair in pending)
                {
                    try
                    {
                        testsData.TestsEditResultStatus((ClashApi.IClashResult)pair.Key, pair.Value);
                        output.updated.Add(pair.Key.DisplayName);
                    }
                    catch (Exception ex)
                    {
                        output.failed[pair.Key.DisplayName] = ex.Message;
                    }
                }

                transaction.Commit();
            }

            return output;
        }

        public static Dictionary<string, int> SetGroupStatuses(
            string testName,
            IEnumerable<string> groupNames,
            string status,
            string comment)
        {
            if (!string.IsNullOrEmpty(comment))
                throw new NotImplementedException(
                    "Добавление комментария к группе коллизий пока не реализовано");

            var requested = new HashSet<string>(StringComparer.Ordinal);
            if (groupNames != null)
            {
                foreach (var name in groupNames)
                    if (!string.IsNullOrWhiteSpace(name)) requested.Add(name.Trim());
            }
            if (requested.Count == 0)
                throw new ArgumentException("Не передано ни одного имени группы");

            var updatedByGroup = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var name in requested) updatedByGroup[name] = 0;

            var testsData = FindManagedTest(testName, out var test);
            var found = requested.ToDictionary(
                name => name,
                name => new List<ClashApi.ClashResultGroup>(),
                StringComparer.Ordinal);
            FindManagedGroups(test, requested, found);

            var missing = found.Where(pair => pair.Value.Count == 0).Select(pair => pair.Key).ToList();
            if (missing.Count > 0)
                throw new ResourceNotFoundException(
                    $"Группы не найдены в тесте '{testName}': {string.Join(", ", missing)}");

            var ambiguous = found.Where(pair => pair.Value.Count > 1).Select(pair => pair.Key).ToList();
            if (ambiguous.Count > 0)
                throw new StateConflictException(
                    $"В тесте '{testName}' найдено несколько групп с одинаковыми именами: " +
                    string.Join(", ", ambiguous));

            var parsed = ParseManagedStatus(status);
            foreach (var pair in found)
            {
                try
                {
                    testsData.TestsEditResultStatus(pair.Value[0], parsed);
                    updatedByGroup[pair.Key] = 1;
                }
                catch (InvalidOperationException ex)
                {
                    throw new StateConflictException(
                        $"Navisworks не разрешил изменить статус группы '{pair.Key}' " +
                        $"в текущем состоянии: {ex.Message}", ex);
                }
                catch (ArgumentException ex)
                {
                    throw new StateConflictException(
                        $"Группа '{pair.Key}' изменилась или больше не принадлежит тесту " +
                        $"'{testName}': {ex.Message}", ex);
                }
            }

            return updatedByGroup;
        }

        public static string ExportReport(string testName, string format)
        {
            // VERIFY (не проверено): вероятный путь — InwOpState10.DriveIOPlugin(
            // pluginName, filePath, options), где pluginName — имя IO-плагина
            // экспорта отчёта Clash Detective (его точное имя не удалось
            // определить статически — это нативный C++/CLI код без видимых
            // управляемых строковых констант). Нужно либо найти имя в
            // документации Navisworks 2020 SDK, либо перебрать
            // State.Plugins() и найти подходящий по ClassName/UserName.
            throw new NotImplementedException(
                "Экспорт отчёта: нужно определить имя IO-плагина экспорта Clash " +
                "Detective и вызвать State.DriveIOPlugin(name, path, options) — " +
                "см. комментарий в коде и README");
        }
    }
}

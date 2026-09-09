using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;          // ModelItem
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
// НАШЛИ (по третьей DLL, Autodesk.Navisworks.ComApi.dll, которую вы
// добавили): настоящий публичный ComApiBridge всё-таки есть в 2020, просто
// лежит в отдельной сборке от Api.dll и Interop.ComApi.dll. State здесь уже
// строго типизирован как InwOpState10 (без каста), и главное — есть
// ComApiBridge.ToModelItem(InwOaPath), которым удобно и надёжно превращать
// путь клэша в обычный managed ModelItem вместо ручного обхода COM Nodes().
using Autodesk.Navisworks.Api.ComApi;

namespace KPLN_NavisMcpBridge
{
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
            throw new InvalidOperationException(
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
            throw new InvalidOperationException($"Тест Clash Detective '{name}' не найден");
        }

        public static List<ClashResultDto> GetResults(string testName, bool includePaths = true, string statusFilter = null, bool includeItemBounds = false, bool includeSizes = false)
        {
            var test = FindTest(testName);
            var list = new List<ClashResultDto>();

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
                if (includePaths || includeItemBounds || includeSizes)
                {
                    item1 = ComApiBridge.ToModelItem(r.Path1);
                    item2 = ComApiBridge.ToModelItem(r.Path2);
                }

                list.Add(new ClashResultDto
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
                    Item1Path = includePaths ? DescribeModelItemPath(item1) : null,
                    Item2Path = includePaths ? DescribeModelItemPath(item2) : null,
                    Comments = comments,
                    Pt1 = DescribePoint(r.Pt1),
                    Pt2 = DescribePoint(r.Pt2),
                    Bound = DescribeBound(r.Bound),
                    Group = DescribeGroup(r)
                });

                if (includeItemBounds)
                {
                    var b1 = DescribeItemBoundingBox(item1);
                    var b2 = DescribeItemBoundingBox(item2);
                    list[list.Count - 1].Item1Bound = b1.Bound;
                    list[list.Count - 1].Item1BoundSource = b1.Source;
                    list[list.Count - 1].Item2Bound = b2.Bound;
                    list[list.Count - 1].Item2BoundSource = b2.Source;
                }

                if (includeSizes)
                {
                    list[list.Count - 1].Item1Size = DescribeItemSize(item1);
                    list[list.Count - 1].Item2Size = DescribeItemSize(item2);
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

        // Имена свойств, в которых лежит сечение. Ищем по ОТОБРАЖАЕМОМУ имени
        // и в русской, и в английской выгрузке; категория не фиксируется —
        // у разных семейств размер лежит то в "Элемент", то в "Тип".
        private static readonly string[] SizeNames = { "Размер", "Size" };
        private static readonly string[] WidthNames = { "Ширина", "Width" };
        private static readonly string[] HeightNames = { "Высота", "Height" };
        private static readonly string[] DiameterNames = { "Диаметр", "Diameter" };

        private static bool NameMatches(string name, string[] candidates)
        {
            foreach (var c in candidates)
                if (string.Equals(name, c, StringComparison.OrdinalIgnoreCase)) return true;
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
        /// Размер элемента из его свойств. Сначала ищем готовую строку
        /// "Размер" ("ø200 мм-ø200 мм", "300x200-300x200"); если её нет —
        /// собираем из Диаметра либо Ширины и Высоты. Если у самого узла
        /// свойств нет (он контейнер), спускаемся к первому потомку с
        /// геометрией — размер лежит там же, где тело.
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
                    if (!child.HasGeometry) continue;
                    var s2 = SizeFromProperties(child);
                    if (!string.IsNullOrEmpty(s2)) return s2;
                    break; // одного тела достаточно, глубже не роемся
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
                foreach (var cat in item.PropertyCategories)
                {
                    foreach (var p in cat.Properties)
                    {
                        var name = p.DisplayName;
                        if (string.IsNullOrEmpty(name)) continue;

                        if (NameMatches(name, SizeNames))
                        {
                            var v = PropertyText(p);
                            if (!string.IsNullOrWhiteSpace(v)) return v; // готовая строка — лучший вариант
                        }
                        else if (diameter == null && NameMatches(name, DiameterNames)) diameter = PropertyText(p);
                        else if (width == null && NameMatches(name, WidthNames)) width = PropertyText(p);
                        else if (height == null && NameMatches(name, HeightNames)) height = PropertyText(p);
                    }
                }
            }
            catch { /* свойства бывают недоступны — не роняем весь запрос */ }

            if (!string.IsNullOrWhiteSpace(diameter)) return "ø" + diameter;
            if (!string.IsNullOrWhiteSpace(width) && !string.IsNullOrWhiteSpace(height))
                return width + "x" + height;
            return null;
        }

        private static string DescribeModelItemPath(Autodesk.Navisworks.Api.ModelItem item)
        {
            if (item == null) return null;
            return string.Join(" / ", item.AncestorsAndSelf.Select(a => a.DisplayName));
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
                throw new InvalidOperationException($"Результат '{resultName}' в тесте '{testName}' не найден");

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

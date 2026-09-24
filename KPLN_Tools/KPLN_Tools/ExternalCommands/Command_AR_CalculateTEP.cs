using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_Tools.ExternalCommands.TepClipper.Clipper;
using CP = KPLN_Tools.ExternalCommands.TepClipper.IntPoint;
using CT = KPLN_Tools.ExternalCommands.TepClipper.ClipType;
using PT = KPLN_Tools.ExternalCommands.TepClipper.PolyType;
using PF = KPLN_Tools.ExternalCommands.TepClipper.PolyFillType;

namespace KPLN_Tools.ExternalCommands
{
    // All Revit access runs in this modal external command's API context.
    // Framework references: WindowsBase, System.Runtime.Serialization, System.Xml.Linq.
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_AR_CalculateTEP : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            if (data.Application.ActiveUIDocument == null) return Result.Cancelled;
            if (data.Application.ActiveUIDocument.Document.IsFamilyDocument)
            { TaskDialog.Show("Расчёт ТЭП", "Откройте проект RVT. Расчёт не выполняется в редакторе семейства."); return Result.Cancelled; }
            try
            {
                var engine = new Engine(data.Application);
                do
                {
                    var window = new KPLN_Tools.Forms.AR_CalculateTEP(engine);
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = data.Application.MainWindowHandle;
                    window.ShowDialog();
                    if (engine.PendingContourAction == null) break;
                    engine.PerformContourAction();
                } while (true);
                engine.NavigateRequested(); return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex) { message = ex.ToString(); return Result.Failed; }
        }
        public const string GeometryVersion = "2026.09.24 / planar-2";
        public const string RulesVersion = "2026.08.26 / 2.0";
        public enum Indicator
        {
            Gns, GnsResidential, GnsLivingPart, GnsNonlivingPart, GnsNonresidential, Footprint,
            Volume, VolumeAbove, VolumeBelow, Gross, GrossAbove, GrossBelow, NnpEmbedded, NnpSeparate,
            Np, NpResidential, NpNonresidential, PublicCalculated, ApartmentsTotal, ApartmentsHeated,
            PublicRooms, ApartmentsCount, ParkingCount, Storeys, Floors
        }
        public class Choice
        {
            public string Key { get; set; }
            public string Label { get; set; }
            public string Description { get; set; }
            public Choice() { }
            public Choice(string key, string label, string description) { Key = key; Label = label; Description = description; }
        }
        private static readonly Dictionary<string, string> MetricLabels = Catalog().ToDictionary(m => m.Key, m => m.Name);
        private static readonly Dictionary<string, string> RoleLabels = Roles().ToDictionary(r => r.Key, r => r.Label);
        private static string MetricLabel(string key) { string label; return key != null && MetricLabels.TryGetValue(key, out label) ? label : key ?? ""; }
        private static string RoleLabel(string key) { string label; return key != null && RoleLabels.TryGetValue(key, out label) ? label : key ?? ""; }
        public class Metric : System.ComponentModel.INotifyPropertyChanged
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string Unit { get; set; }
            public string Description { get; set; }
            public bool Enabled { get; set; }
            public string Geometry { get; set; }
            public string AreaScheme { get; set; }
            private string contourMode = "current";
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            public string ContourMode { get { return contourMode; } set { if (contourMode == value) return; contourMode = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(ContourMode))); } }
            public string WallSelectionMode { get; set; } = "exterior";
            public string WallParameter { get; set; }
            public string WallValue { get; set; }
            public string WallBoundary { get; set; } = "auto";
            public string WallCutHeight { get; set; } = "1";
            public List<string> SelectedWalls { get; set; } = new List<string>();
            public string ContourExclusions { get; set; } = "normative";
            public string VolumeMaskHeight { get; set; } = "actual";
        }
        public static bool SupportsContours(string key)
        { return new[] { "Gns", "GnsResidential", "GnsNonresidential", "Gross", "GrossAbove", "GrossBelow", "Np", "NpResidential", "NpNonresidential", "Volume", "VolumeAbove", "VolumeBelow", "Footprint" }.Contains(key); }
        public class ContourSketch
        {
            public string Metric { get; set; }
            public string ElementUniqueId { get; set; }
            public string LevelKey { get; set; }
            public string LevelName { get; set; }
            public string Kind { get; set; } = "base";
            public string Building { get; set; }
            public string Section { get; set; } = "";
            public string Profile { get; set; } = "residential";
            public string BuildingClass { get; set; } = "residential";
            public string Height { get; set; }
            public string BottomOffset { get; set; } = "0";
            public string ElementLabel { get; set; }
        }
        public static List<Metric> Catalog()
        {
            string[] names = { "1. Суммарная поэтажная площадь всего", "2. СПП жилых зданий", "2а. Жилая часть СПП жилых зданий", "2б. Нежилая часть СПП жилых зданий", "3. СПП нежилых зданий", "4. Площадь застройки", "5. Строительный объём", "5а. Строительный объём выше 0.000", "5б. Строительный объём ниже 0.000", "6. Общая площадь объекта", "6а. Общая площадь наземной части", "6б. Общая площадь подземной части", "7. ННП встроенно-пристроенная", "8. ННП встроенно-пристроенная отдельно стоящая", "9. Наземная площадь всего", "9а. Наземная площадь жилых зданий", "9б. Наземная площадь нежилых зданий", "10. Расчётная площадь общественного здания", "11. Общая площадь квартир с летними помещениями", "12. Площадь квартир без летних помещений", "13. Площадь помещений общественного назначения", "14. Количество квартир", "15. Машино-места подземного паркинга", "16. Этажность", "17. Количество этажей" };
            string[] descriptions = {
                "Сумма наземных этажей по наружному обмеру стен. Состав проёмов и шахт зависит от выбранной старой/новой методики.",
                "ГНС только зданий, отнесённых к жилым; включает жилую и нежилую части.", "ГНС жилых зданий: квартиры, входные группы жилья и обслуживающая жильё часть, заданная классификацией.", "ГНС жилых зданий: нежилые функции; граница с жилой частью проходит по оси разделяющих стен.", "ГНС зданий, отнесённых к нежилым.",
                "Объединение сечения внешнего объёма на отметке земли, учитываемых выступов и выступающего подземного контура. Консоли учитываются при высоте менее 4,5 м.",
                "Объединённый замкнутый внешний строительный объём с вычитанием нормативных исключений. Сумма частей выше и ниже 0.000.", "Часть строительного объёма выше выбранной плоскости 0.000.", "Часть строительного объёма ниже выбранной плоскости 0.000.",
                "Сумма этажей по внутренней поверхности наружных стен, включая внутренние стены. Исключения определяет профиль здания.", "Общая площадь только наземных этажей.", "Общая площадь только подземных этажей.",
                "Наземная нежилая площадь с признаком встроенно-пристроенной части и без признака отдельно стоящего объекта.", "Наземная нежилая площадь с обоими признаками: встроенно-пристроенная часть и отдельно стоящий объект.",
                "Общая площадь наземных этажей всех зданий.", "Наземная площадь только жилых зданий.", "Наземная площадь только нежилых зданий.",
                "Площадь общественных помещений без коридоров, тамбуров, переходов, лестниц, пандусов, шахт и инженерных помещений; высота под наклонными поверхностями не ниже 1,5 м.",
                "Отапливаемые помещения квартир плюс учитываемые летние помещения с заданными коэффициентами. Французские балконы исключаются.", "Только отапливаемые помещения квартир и антресоли; летние и неотапливаемые помещения исключаются.", "Общественные помещения по внутренним отделанным поверхностям, включая относящиеся к ним наружные части по профилю.",
                "Уникальные ID квартир в пределах экземпляра связи, корпуса и секции либо экземпляры семейств по выбранному режиму.", "Только машино-места подземной части; режим по ID выявляет дубли, режим по экземплярам считает элементы.",
                "Учитываемые наземные этажи по корпусам и секциям. Итог проекта - максимум, а не сумма этажностей корпусов.", "Все учитываемые наземные и подземные этажи по корпусам и секциям. Итог проекта - максимум." };
            return Enum.GetValues(typeof(Indicator)).Cast<Indicator>().Select((v, i) => new Metric { Key = v.ToString(), Name = names[i], Unit = i >= 23 ? "этажей" : i >= 21 ? "шт." : i >= 6 && i <= 8 ? "м³" : "м²", Enabled = true, Description = descriptions[i] }).ToList();
        }
        public static List<Choice> Choices(string kind)
        {
            if (kind == "contour") return new List<Choice>{
                new Choice("current","Как сейчас (по умолчанию)","Сохраняет расчёт по помещениям, зонам и геометрии модели. Для объёма используются пространственные элементы и конструкции либо заданная оболочка. Площади и распознанные вертикальные участки вычитаются по плоским контурам; сложные тела используют Revit."),
                new Choice("walls","По наружным стенам","Строит замкнутые контуры из выбранных вертикальных базовых стен на каждом уровне. Не требует объединения всех конструкций здания. Разрывы, ответвления и неподдерживаемые стены вызывают ошибку."),
                new Choice("manual","Чертим контур","Использует закреплённые цветовые области активной модели. Граница области непосредственно задаёт границу обмера; толщина стен к ней не прибавляется. Для объёма строятся вертикальные призмы по этажам.")};
            if (kind == "wall-selection") return new List<Choice>{
                new Choice("exterior","Функция стены: Наружная","Берёт стены с функцией типа «Наружная» из включённых источников. Внутренний двор должен быть замкнут отдельным кольцом стен."),
                new Choice("parameter","По параметру и значению","Берёт стены, у которых параметр экземпляра или типа равен заданному значению без учёта регистра."),
                new Choice("selection","Выбранные вручную","Использует закреплённый выбор стен активной модели. Выбор стен связей выполняется режимом функции или параметра.")};
            if (kind == "wall-boundary") return new List<Choice>{
                new Choice("auto","По показателю и методике","Для ГНС: старая методика - наружная поверхность, новая - наружная граница ядра. Общая и наземная площадь - внутренняя поверхность наружных стен. Объём и застройка - наружная поверхность."),
                new Choice("exterior","Наружная поверхность","Обмер по наружной боковой поверхности стен с отделкой."),
                new Choice("core","Наружная граница ядра","Из толщины наружной стороны стены убирает наружные слои до ядра."),
                new Choice("interior","Внутренняя поверхность","Обмер по внутренней боковой поверхности наружных стен.")};
            if (kind == "contour-exclusions") return new List<Choice>{
                new Choice("normative","Методика и правила показателя","Вырезает помещения и элементы, исключённые методикой, профилем, параметром включения или правилами. Неопределённое назначение помещения даёт ошибку. Ручные вырезы применяются дополнительно."),
                new Choice("rules","Только явные исключения","Вырезает объекты с явным исключением правилом, корректировкой или параметром включения, а также ручные вырезы. Нормативные исключения по назначению автоматически не применяются; состав вырезов задаёт пользователь.")};
            if (kind == "mask-height") return new List<Choice>{
                new Choice("actual","По реальной геометрии","Из объёма вычитается фактический объём помещения или элемента в его высотном диапазоне. Для зоны без объёма требуется режим на высоту участка либо ручной вырез."),
                new Choice("floor","На всю высоту участка","Проекция исключаемого объекта вычитается на всю высоту расчётного участка его этажа. Это сквозной вырез, который может быть больше реального объёма объекта.")};
            if (kind == "sketch-kind") return new List<Choice> { new Choice("base", "Основной контур", "Задаёт исходную площадь этажа в ручном режиме."), new Choice("exclude", "Вырез", "Вычитается из основного контура. Для объёма учитываются смещение низа и высота выреза.") };
            if (kind == "diagnostics") return new List<Choice> { new Choice("full", "Подробно", "Показывает каждое сообщение с исходным элементом. Экспорт всегда содержит полный журнал."), new Choice("brief", "Кратко", "Объединяет одинаковые сообщения по показателю, источнику и корпусу; показывает число повторений. Для перехода к каждому объекту выберите подробный режим.") };
            if (kind == "method") return new List<Choice> {
                new Choice("new","Новая методика","Проёмы любой площади и шахты учитываются на одном нижнем этаже. Террасы и эксплуатируемая кровля исключаются из ГНС. Автоматическая обводка базовой стены проходит по наружной границе ядра, без внешних отделочных слоёв."),
                new Choice("old","Старая методика","Для проёмов правило одного этажа действует при площади более 36 м². Для шахт отдельное автоматическое исключение не применяется; террасы задаются правилами проекта. Автоматическая обводка использует полную толщину наружной стены.") };
            if (kind == "profile") return new List<Choice> {
                new Choice("residential","Жилое здание до 75 м","Из общей площади исключаются технические пространства, подполья и чердаки. Переход между корпусами делится поровну. Цоколь входит в этажность при превышении верха перекрытия над землёй не менее 2 м."),
                new Choice("public","Общественное здание до 50 м","Технические помещения входят в общую площадь; техническое пространство ниже 1,8 м и кровельные надстройки при выполнении условий исключаются. Цоколь входит в этажность. Для расчётной площади исключаются коммуникационные и инженерные помещения."),
                new Choice("high-residential","Высотное жилое здание","Правила жилого здания; дополнительно исключаются технические пространства, занимающие часть этажа. Верхняя надстройка менее 8 м² и высотой менее 2,5 м не образует последний этаж."),
                new Choice("high-public","Высотное общественное здание","Площади и объём по общественному профилю; площадь застройки и этажность по жилому. Техническое пространство части этажа исключается."),
                new Choice("high-mixed","Высотный многофункциональный комплекс","Нужно задать профиль функциональных частей в таблице корпусов. Без этого неоднозначные показатели отмечаются как неполные; автоматическая подмена смешанного профиля жилым не выполняется."),
                new Choice("by-building","Определять по корпусам","Для каждого корпуса используется профиль из таблицы соответствий. Корпуса без профиля попадут в ошибки."),
                new Choice("by-source","Определять по связям","Для каждого экземпляра связи используется его профиль из таблицы источников. Одна связь может иметь общий профиль для нескольких корпусов.") };
            if (kind == "source") return new List<Choice> { new Choice("include", "Включать", "Источник участвует в числах и проверочной графике."), new Choice("exclude", "Исключить", "Источник не участвует в расчёте и графике."), new Choice("reference", "Только графическая проверка", "Геометрия показывается как справочная и не меняет числовые итоги.") };
            if (kind == "group") return new List<Choice> { new Choice("links", "По связям Revit", "Каждый экземпляр связи - отдельный корпус; одинаковые файлы с разными размещениями не смешиваются."), new Choice("parameter", "По параметру корпуса", "Объединяет элементы по параметру корпуса, в том числе из нескольких связей."), new Choice("worksets", "По рабочим наборам", "Корпус определяется рабочим набором элемента; для несотрудничающих моделей нужна другая схема."), new Choice("mapping", "По таблице соответствий", "Корпуса определяются пользовательскими правилами по источнику и значению параметра."), new Choice("selection", "Ручной выбор", "Считает только элементы или экземпляры связей, выделенные перед запуском команды, с указанным именем корпуса.") };
            if (kind == "geometry") return new List<Choice> { new Choice("auto", "Автоматически", "На каждом уровне: зоны выбранной схемы, иначе помещения, иначе пространства. Параллельные представления одной площади не складываются."), new Choice("rooms", "Помещения", "Расчёт по Rooms; Spaces и Areas не добавляются поверх них."), new Choice("spaces", "Пространства", "Расчёт по MEP Spaces; Rooms и Areas исключаются как альтернативные представления."), new Choice("areas", "Зоны", "Используются границы выбранной схемы Areas. Для ГНС зона должна быть построена по внешнему контуру; для общей площади - по внутреннему.") };
            if (kind == "graphics") return new List<Choice> { new Choice("replace", "Обновить проверочные данные", "Заменяет только помеченную плагином графику. При ошибке транзакция откатывается."), new Choice("new", "Создать новый набор", "Сохраняет предыдущие наборы; новые виды и элементы получают идентификатор текущего расчёта."), new Choice("none", "Не создавать графику", "Выполняется только расчёт. Существующие проверочные элементы не изменяются.") };
            if (kind == "datum") return new List<Choice> { new Choice("manual", "Ввести отметку", "Используется заданное число в метрах относительно внутренних координат проекта."), new Choice("level", "Уровень", "Используется отметка выбранного уровня с учётом размещения связи."), new Choice("parameter", "Параметр проекта", "Используется параметр активной модели типа длина либо числовой текст в метрах."), new Choice("terrain", "По поверхности", "Для горизонтальной топографии / топотела использует среднюю отметку вершин. Если перепад больше 1 см, требует ручную отметку и уточнение наземности по секциям; единая отметка для наклонного участка не подставляется.") };
            if (kind == "count") return new List<Choice> { new Choice("id", "По уникальным ID", "Один ID в пределах источника, корпуса и секции считается один раз. Пустые ID и дубли семейств выводятся в ошибки."), new Choice("instances", "По экземплярам семейств", "Каждый распознанный экземпляр семейства считается единицей; помещения не считаются экземплярами квартиры.") };
            if (kind == "level") return new List<Choice> { new Choice("normal", "Обычный этаж", "Определяется по отметке пола относительно земли; принудительное отнесение доступно в соседней колонке."), new Choice("basement", "Цокольный", "Для ГНС и жилой этажности требуется верх перекрытия не ниже средней земли +2 м; для общественной этажности учитывается как цоколь."), new Choice("technical", "Технический этаж", "Учитывается, если представляет этаж, а не исключаемое техническое пространство."), new Choice("attic", "Чердак", "Исключается из площади здания и этажности."), new Choice("void", "Подполье / техническое пространство", "Исключается либо проверяется по высоте согласно профилю здания."), new Choice("roof", "Кровельная надстройка", "Учитываются назначение, высота и отношение площади к кровле; без этих данных будет сообщение о неполноте."), new Choice("exclude", "Не является этажом", "Служебный уровень исключается из расчёта площадей и этажности.") };
            if (kind == "above") return new List<Choice> { new Choice("auto", "Автоматически", "Применяет отметку земли и правило цоколя; на уклоне используйте отдельные секции."), new Choice("above", "Наземный", "Принудительно относит уровень к наземной части; фиксируется в настройках отчёта."), new Choice("below", "Подземный", "Принудительно относит уровень к подземной части; фиксируется в настройках отчёта.") };
            if (kind == "action") return new List<Choice> { new Choice("classify", "Классифицировать", "Назначает функциональный тип без изменения параметров Revit."), new Choice("include", "Включить", "Явно включает элемент в выбранный показатель; решение отражается в отчёте."), new Choice("exclude", "Исключить", "Исключает элемент и при наличии контура вычитает его площадь из других включённых контуров.") };
            if (kind == "operation") return new List<Choice> { new Choice("equals", "Равно", "Точное совпадение текста без учёта регистра и пробелов по краям."), new Choice("contains", "Содержит", "Совпадение части текста; может распознать несколько значений."), new Choice("filled", "Заполнено", "Параметр существует и содержит непустое значение."), new Choice("empty", "Пусто", "Параметр существует, но не заполнен. Отсутствующий параметр не считается пустым.") };
            if (kind == "correction") return new List<Choice> { new Choice("include", "Включить объект", "Добавляет объект в показатель с записью автора, даты и причины."), new Choice("exclude", "Исключить объект", "Убирает объект из показателя с записью автора, даты и причины."), new Choice("classify", "Изменить классификацию", "Применяет указанную роль только для расчёта, исходная модель не изменяется."), new Choice("building", "Уточнить корпус", "Заменяет корпус объекта только в расчёте."), new Choice("delta", "Корректирующее значение", "Прибавляет число в единицах показателя. Геометрия для числовой корректировки не выдумывается; в отчёте будет отдельная строка.") };
            if (kind == "part") return new List<Choice> { new Choice("auto", "По назначению", "Жилая часть определяется ролью помещения; техническую обслуживающую часть задайте явно."), new Choice("residential", "Жилая часть", "Относит помещение или обслуживающий объект к жилой части здания."), new Choice("nonresidential", "Нежилая часть", "Относит помещение или объект к нежилой части здания.") };
            return Roles();
        }
        public static List<Choice> Roles()
        {
            string[] keys = { "unknown", "heated", "auxiliary", "public", "residential-common", "nonresidential-common", "technical-room", "technical-space", "technical-void", "attic", "roof-exit", "roof-vent", "roof-used", "balcony", "loggia", "terrace", "veranda", "cold-storage", "tambour", "external-tambour", "entrance", "stair", "open-stair", "ramp", "corridor", "transition", "multilight", "shaft", "opening", "stair-gap", "parking", "niche", "arch", "under-stair", "structure", "envelope", "footprint", "underground-footprint", "canopy", "porch", "pit", "passage", "decoration", "stove", "mezzanine", "french-balcony" };
            string[] labels = { "Не определено", "Жилое помещение квартиры", "Отапливаемое вспомогательное помещение квартиры", "Общественное помещение", "МОП жилой части", "МОП нежилой части", "Техническое помещение", "Техническое пространство", "Подполье", "Чердак", "Выход на кровлю", "Венткамера на кровле", "Эксплуатируемая кровля", "Балкон", "Лоджия", "Терраса", "Веранда", "Холодная кладовая", "Тамбур", "Наружный тамбур", "Входная группа жилья", "Лестница / лестничная клетка", "Наружная открытая лестница", "Пандус", "Коридор", "Переход между корпусами", "Многосветное пространство", "Шахта", "Проём", "Лестничный просвет", "Машино-место", "Ниша", "Арочный проём", "Под внутриквартирной лестницей", "Конструкция", "Замкнутый внешний расчётный объём", "Контур застройки по земле", "Подземный контур", "Козырёк / навес / консоль", "Крыльцо / входная площадка", "Приямок", "Проезд / пространство под зданием", "Декоративный элемент", "Отопительная печь", "Антресоль", "Французский балкон" };
            var notes = new Dictionary<string, string>{
                {"unknown","Объект не получает предполагаемое назначение. Для показателей, которым нужна классификация, выводится ошибка."},
                {"heated","Входит в отапливаемую площадь квартиры при наличии ID; в ГНС относится к жилой части."},
                {"auxiliary","Входит в площадь квартиры без летних помещений и с ними, с коэффициентом 1."},
                {"public","Входит в общественные площади и расчётную площадь; для помещения применяется общественный профиль, включая пороги высоты."},
                {"residential-common","Входит в жилую часть ГНС, но не в площадь квартир без явного ID и квартирного назначения."},
                {"nonresidential-common","Входит в нежилую часть ГНС; для площади квартир не используется."},
                {"technical-room","Включается в общую площадь по профилю, исключается из расчётной общественной площади."},
                {"technical-space","В жилом здании исключается из общей площади. В общественном пространстве ниже 1,8 м требуется признак прохода обслуживания."},
                {"technical-void","В жилом здании исключается из площади и этажности. Для общественного применяется порог 1,8 м. Сам по себе не означает исключение из строительного объёма."},
                {"attic","Исключается из общей площади. Строительный объём включает чердак в пределах наружной оболочки."},
                {"roof-exit","Исключается из общей площади и проверяется отдельно при определении последнего этажа высотного здания."},
                {"roof-vent","Для общественного здания учитывается суммарная доля надстроек: менее 15% кровли исключается из общей площади."},
                {"roof-used","Входит в общую площадь; новая методика исключает эксплуатируемую кровлю из ГНС."},
                {"balcony","Площадь квартиры с летними помещениями: коэффициент 0,3. Из отапливаемой площади и строительного объёма исключается."},
                {"loggia","Площадь квартиры с летними помещениями: коэффициент 0,5. Из площади без летних помещений исключается."},
                {"terrace","Квартирная площадь: коэффициент 0,3; новая методика ГНС и строительный объём исключают террасу."},
                {"veranda","Входит в квартирную площадь с летними помещениями с коэффициентом 1; в отапливаемую площадь не входит."},
                {"cold-storage","Квартирная площадь с летними помещениями: коэффициент 1; в площадь отапливаемых помещений не входит."},
                {"tambour","Неотапливаемый квартирный тамбур входит только в общую площадь квартиры. Из общей площади жилого здания и расчётной общественной площади исключается."},
                {"external-tambour","В общественном профиле входит в общую площадь; при нежилой принадлежности входит в площадь общественного помещения."},
                {"entrance","Входная группа жилья относится к жилой части ГНС. Для площади застройки учитывается как выступающая часть жилого здания."},
                {"stair","Включается поэтажная проекция площадок и ступеней; из расчётной площади общественного здания исключается."},
                {"open-stair","Наружная открытая лестница исключается из общей площади и ГНС; для включения в застройку задайте контур крыльца / входных ступеней."},
                {"ramp","Из общей площади жилого здания и расчётной общественной площади исключается; общая площадь общественного здания включает внутренние рампы."},
                {"corridor","Входит в общую площадь, но не в расчётную площадь общественного здания. Обслуживаемую жилую / нежилую часть задайте явно."},
                {"transition","В общей площади жилых зданий делится поровну между корпусами из параметра перехода. Из расчётной общественной площади исключается."},
                {"multilight","Учитывается на нижнем этаже; верхние контуры с тем же ID вертикального пространства вычитаются."},
                {"shaft","Новая методика ГНС учитывает шахту один раз на нижнем этаже; старая не применяет это отдельное исключение. Общая площадь учитывает на одном этаже."},
                {"opening","Новая методика ГНС: один нижний этаж для любой площади. Старая: правило применяется только при площади более 36 м²."},
                {"stair-gap","Правило одного этажа применяется при ширине просвета более ширины марша либо 1,5 м. Нужны параметры ширин и ID вертикального пространства."},
                {"parking","Для количества учитывается только подземная часть; ID выявляет повторные элементы одного места."},
                {"niche","В жилой площади учитывается при высоте не менее 2 м. Нужен параметр высоты."},
                {"arch","В жилой площади учитывается при ширине не менее 2 м. Нужен параметр ширины."},
                {"under-stair","Учитывается только часть пола с высотой более 1,6 м; для Rooms/Spaces контур обрезается по сечению фактического объёма."},
                {"structure","Участвует в восстановлении внешнего объёма. Площадь материала конструкции не суммируется с площадью помещений."},
                {"envelope","Авторитетное замкнутое тело строительного объёма, включая ограждения. При его наличии восстановление по Rooms/Spaces не используется."},
                {"footprint","Явный внешний контур у планировочной отметки земли. Используется вместо автоматического сечения объёма."},
                {"underground-footprint","Проекция выступающей подземной части объединяется с наземным контуром застройки без повторного учёта пересечений."},
                {"canopy","Из строительного объёма исключается. В застройке консоль учитывается при высоте менее 4,5 м над землёй."},
                {"porch","Включает крыльцо, входные площадки и ступени в площадь застройки, но не в общую площадь этажей."},
                {"pit","Приямок включается в площадь застройки; в поэтажные площади помещений не включается."},
                {"passage","Проезд и пространство под зданием входят в застройку, но вычитаются из строительного объёма."},
                {"decoration","Не включается в строительный объём и общую площадь; исключаемый контур вычитается из учитываемой площади."},
                {"stove","Контур отопительной печи вычитается из площади пола. Для декоративного камина требуется отдельное решение."},
                {"mezzanine","Входит в площадь квартиры. Для общей площади многоэтажного общественного здания применяется порог более 40% площади этажа; в одноэтажном учитывается полностью."},
                {"french-balcony","Не учитывается как летнее помещение квартиры и не добавляется к общей площади."}
            };
            var result = keys.Select((k, i) => new Choice(k, labels[i], notes[k])).ToList();
            result.Add(new Choice("apartment-family", "Семейство целой квартиры", "В режиме подсчёта по экземплярам один распознанный экземпляр означает одну квартиру. Мебель и другие семейства с тем же ID квартиры не считаются квартирами."));
            result.Add(new Choice("ventilated-void", "Проветриваемое подполье", "Исключается из строительного объёма и общей площади. Для общественного здания используйте эту роль для подполья на многолетнемерзлых грунтах."));
            result.Add(new Choice("soil-filled", "Засыпанное землёй пространство", "Исключается из общей площади и расчётного объёма, в отличие от полноценного помещения подземного этажа."));
            result.Add(new Choice("engineering-shaft", "Шахта вертикальных инженерных коммуникаций", "В жилой общей площади не учитывается. В новой методике ГНС учитывается один раз на нижнем этаже по общему ID, как шахта."));
            return result;
        }
        public class ParameterMap { public string Key { get; set; } public string Title { get; set; } public string Name { get; set; } public string Description { get; set; } }
        public class Rule
        {
            public bool Enabled { get; set; } = true;
            public string Metric { get; set; } = "all";
            public string Category { get; set; }
            public string Parameter { get; set; } = "@Name";
            public string Operation { get; set; } = "contains";
            public string Value { get; set; }
            public string Role { get; set; } = "unknown";
            public string Part { get; set; } = "auto";
            public string Action { get; set; } = "classify";
            public double? Coefficient { get; set; }
        }
        public class SourceOption { public string Key { get; set; } public string Mode { get; set; } public string Profile { get; set; } }
        public class BuildingMap
        {
            public string Source { get; set; } = "*";
            public string MatchValue { get; set; } = "*";
            public string Building { get; set; }
            public string Profile { get; set; } = "residential";
            public string Class { get; set; } = "residential";
            public bool Include { get; set; } = true;
        }
        public class LevelSetting
        {
            public string Key { get; set; }
            public string Source { get; set; }
            public string Name { get; set; }
            public string Building { get; set; }
            public string Section { get; set; }
            public double Elevation { get; set; }
            public bool Include { get; set; } = true;
            public string Kind { get; set; } = "normal"; public string Above { get; set; } = "auto";
            public string TopSlab { get; set; }
            public string Height { get; set; }
            public string RoofRatio { get; set; }
            public string RoofArea { get; set; }
            public double ElevationMeters { get { return Elevation * .3048; } }
            public string DisplayName { get { return Source + " / " + Name + " (" + ElevationMeters.ToString("0.000") + " м)"; } }
        }
        public class Correction
        {
            public string Metric { get; set; }
            public string Action { get; set; }
            public string Source { get; set; }
            public string Element { get; set; }
            public string Value { get; set; }
            public string Reason { get; set; }
            public string Author { get; set; }
            public string Date { get; set; }
            public string Previous { get; set; }
        }
        public class Settings : System.ComponentModel.INotifyPropertyChanged
        {
            public ObservableCollection<ContourSketch> Contours { get; set; } = new ObservableCollection<ContourSketch>();
            public int Version { get; set; } = 2;
            public string Method { get; set; } = "new";
            public string Profile { get; set; } = "residential";
            public string Grouping { get; set; } = "links";
            public string ManualBuilding { get; set; } = "Корпус 1";
            public string Geometry { get; set; } = "auto";
            public string Phase { get; set; }
            public string AreaScheme { get; set; }
            public string ZeroMode { get; set; } = "manual"; public string ZeroValue { get; set; } = "0";
            public string ZeroLevel { get; set; }
            public string ZeroParameter { get; set; }
            public string GroundMode { get; set; } = "manual"; public string GroundValue { get; set; } = "0";
            public string GroundLevel { get; set; }
            public string GroundParameter { get; set; }
            public string ApartmentMode { get; set; } = "id"; public string ParkingMode { get; set; } = "id";
            private bool createViews;
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            public bool CreateViews { get { return createViews; } set { if (createViews == value) return; createViews = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(CreateViews))); } }
            public string Graphics { get; set; } = "replace"; public bool Create3D { get; set; } = true;
            public bool CreateSchedule { get; set; } = false;
            public string Prefix { get; set; } = "ТЭП_";
            public int Decimals { get; set; } = 2;
            public string IssueDetail { get; set; } = "full";
            public List<Metric> Metrics { get; set; } = Catalog();
            public List<ParameterMap> Parameters { get; set; } = DefaultParameters();
            public ObservableCollection<Rule> Rules { get; set; } = new ObservableCollection<Rule>();
            public ObservableCollection<BuildingMap> Buildings { get; set; } = new ObservableCollection<BuildingMap>();
            public ObservableCollection<LevelSetting> Levels { get; set; } = new ObservableCollection<LevelSetting>();
            public ObservableCollection<Correction> Corrections { get; set; } = new ObservableCollection<Correction>();
            public List<SourceOption> Sources { get; set; } = new List<SourceOption>();
            public string IncludedColor { get; set; } = "#347EC4";
            public string ExcludedColor { get; set; } = "#CC5555";
            public string ManualColor { get; set; } = "#A56ED3";
            public string PublicColor { get; set; } = "#C59043";
            public string SummerColor { get; set; } = "#36A38E";
            public string TechnicalColor { get; set; } = "#8993A0";
            public string ErrorColor { get; set; } = "#E24040";
            public string PlanTemplate { get; set; }
            public string View3DTemplate { get; set; }
            public static List<ParameterMap> DefaultParameters()
            {
                string[] keys = { "building", "profile", "function", "part", "apartment", "section", "parking", "embedded", "standalone", "vertical", "coefficient", "height", "slope", "width", "roof-ratio", "partial-floor", "transition", "include", "mezzanine-ratio" };
                string[] labels = { "Корпус", "Тип здания / профиль", "Назначение помещения", "Жилая / нежилая часть", "Номер / ID квартиры", "Секция", "ID машино-места", "Признак встроенно-пристроенной части", "Признак отдельно стоящего объекта", "ID вертикального пространства", "Коэффициент летнего помещения", "Высота в свету", "Угол наклона потолка, градусы", "Ширина проёма / просвета", "Доля площади надстройки от кровли", "Техническое пространство занимает часть этажа", "Корпуса перехода через ;", "Признак включения в ТЭП", "Доля площади антресоли от этажа" };
                var result = keys.Select((k, i) => new ParameterMap { Key = k, Title = labels[i], Name = "", Description = "Параметр экземпляра или типа: " + labels[i] + ". Длины Revit переводятся в метры; текстовые длины задаются в метрах, доли в диапазоне 0..1. Для признаков: 1/0, да/нет." }).ToList();
                result.Add(new ParameterMap { Key = "service-access", Title = "Нужен проход для обслуживания коммуникаций", Name = "", Description = "Для общественного технического пространства ниже 1,8 м: да/1 - включать, нет/0 - исключать. Отсутствие признака вызывает ошибку." });
                result.Add(new ParameterMap { Key = "stair-width", Title = "Ширина лестничного марша", Name = "", Description = "Длина Revit или текст в метрах. Для лестничного просвета проверяются ширина марша и порог 1,5 м." });
                result.Add(new ParameterMap { Key = "level", Title = "Расчётный уровень объекта", Name = "", Description = "Параметр-ссылка на уровень либо его точное имя. При заполнении уточняет встроенный LevelId; полезен для расчётных семейств и DirectShape." });
                return result;
            }
            public string Parameter(string key) { var p = Parameters.FirstOrDefault(x => x.Key == key); return p == null ? "" : p.Name; }
        }
        public class Source
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string Mode { get; set; } = "include";
            public string Profile { get; set; } = "residential"; public bool Loaded { get { return Document != null; } }
            public string LoadError { get; set; }
            internal Document Document; internal Transform Transform; internal ElementId RootLink;
            internal List<Element> Elements = new List<Element>();
        }
        public class Issue
        {
            public string Code { get; set; }
            public string Severity { get; set; }
            public string Message { get; set; }
            public string Metric { get; set; }
            public string Source { get; set; }
            public string Building { get; set; }
            public string Element { get; set; }
            public string Action { get; set; }
            public string MetricName { get { return string.IsNullOrEmpty(Metric) ? "Все показатели" : MetricLabel(Metric); } }
        }
        public class Detail
        {
            public string Metric { get; set; }
            public string SourceKey { get; set; }
            public string Source { get; set; }
            public string Building { get; set; }
            public string Section { get; set; }
            public string Level { get; set; }
            public double Elevation { get; set; }
            public string Element { get; set; }
            public string UniqueId { get; set; }
            public string Purpose { get; set; }
            public string Apartment { get; set; }
            public string Profile { get; set; }
            public string Method { get; set; }
            public string Unit { get; set; }
            public double Raw { get; set; }
            public double Factor { get; set; } = 1;
            public double Value { get; set; }
            public bool Excluded { get; set; }
            public bool Manual { get; set; }
            public string Reason { get; set; }
            public string ViewId { get; set; }
            public bool HasError { get; set; }
            public string MetricName { get { return MetricLabel(Metric); } }
            public string PurposeName { get { return RoleLabel(Purpose); } }
            internal Solid Shape; internal bool VolumeShape;
        }
        public class Summary
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public double Value { get; set; }
            public string Unit { get; set; }
            public string Method { get; set; }
            public string Status { get; set; }
            public string Comment { get; set; }
        }
        public class Run
        {
            public string Id { get; set; } = Guid.NewGuid().ToString("N"); public string Date { get; set; } = DateTime.Now.ToString("s");
            public string GeometryVersion { get; set; } = Command_AR_CalculateTEP.GeometryVersion;
            public bool? CreateViews { get; set; }
            public string Author { get; set; }
            public string Method { get; set; }
            public string Version { get; set; } = RulesVersion;
            public string Configuration { get; set; }
            public List<Summary> Summary { get; set; } = new List<Summary>();
            public List<Detail> Details { get; set; } = new List<Detail>(); public List<Issue> Issues { get; set; } = new List<Issue>();
            public List<Correction> Corrections { get; set; } = new List<Correction>();
            public void Issue(string code, string severity, string message, string metric = "", string source = "", string building = "", string element = "", string action = "Проверьте настройки и исправьте исходные данные.")
            { Issues.Add(new Issue { Code = code, Severity = severity, Message = message, Metric = metric, Source = source, Building = building, Element = element, Action = action }); }
            public string TextReport(int decimals)
            {
                var b = new StringBuilder("Расчёт ТЭП | " + Date + " | " + Method + " | правила " + Version + (string.IsNullOrEmpty(GeometryVersion) ? "" : " | геометрия " + GeometryVersion) + (CreateViews.HasValue ? " | проверочные виды: " + (CreateViews.Value ? "включены" : "выключены") : "") + "\n");
                foreach (var s in Summary)
                {
                    b.AppendLine(); b.AppendLine(s.Name + ": " + Math.Round(s.Value, decimals, MidpointRounding.AwayFromZero).ToString("N" + decimals) + " " + s.Unit + " | " + s.Status);
                    foreach (var g in Details.Where(d => d.Metric == s.Key && !d.Excluded).GroupBy(d => new { d.Building, d.Section }).OrderBy(g => g.Key.Building))
                    {
                        b.AppendLine("  " + g.Key.Building + (string.IsNullOrWhiteSpace(g.Key.Section) ? "" : " / секция " + g.Key.Section));
                        foreach (var floor in g.GroupBy(d => Math.Round(d.Elevation, 6)).OrderBy(x => x.Key))
                            b.AppendLine("    " + string.Join(" / ", floor.Select(d => d.Level).Distinct()) + " (" + (floor.Key * .3048).ToString("+0.000;-0.000;0.000") + " м): " + Math.Round(floor.Sum(d => d.Value), decimals, MidpointRounding.AwayFromZero).ToString("N" + decimals) + " " + s.Unit);
                    }
                    if (!string.IsNullOrEmpty(s.Comment)) b.AppendLine("  " + s.Comment);
                }
                return b.ToString();
            }
        }
        internal class Record
        {
            internal Source Source; internal Element Element; internal LevelSetting Level;
            internal string Building, Section, Profile, BuildingClass, Role, Part, Apartment, Vertical;
            internal double? Factor; internal bool Manual; internal bool? Override; internal bool SingleStorey; internal bool BuildingIncluded = true;
            internal string Key { get { return Source.Key + "/" + Element.UniqueId; } }
            internal double Z { get { return Level == null ? 0 : Level.Elevation; } }
            internal Record Copy() { return (Record)MemberwiseClone(); }
        }
        // Immutable numeric profiles. XY is in Revit feet; integer grid is 0.001 mm.
        internal sealed class PlanarRegion
        {
            internal const double Scale = 304800.0;
            internal readonly List<List<CP>> Paths;
            internal static readonly PlanarRegion Empty = new PlanarRegion(new List<List<CP>>());
            internal PlanarRegion(List<List<CP>> paths) { Paths = paths; }
            internal bool IsEmpty { get { return Paths.Count == 0; } }
            internal static long Coordinate(double feet)
            {
                double value = feet * Scale;
                if (double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) > 1e13) throw new InvalidOperationException("Координата контура вне безопасного диапазона плоского расчёта.");
                return checked((long)Math.Round(value, MidpointRounding.AwayFromZero));
            }
            private static double RingArea(List<CP> path)
            {
                if (path.Count < 3) return 0; double sum = 0; var origin = path[0];
                for (int i = 1; i + 1 < path.Count; i++)
                    sum += ((double)path[i].X - origin.X) * ((double)path[i + 1].Y - origin.Y) - ((double)path[i + 1].X - origin.X) * ((double)path[i].Y - origin.Y);
                return sum / (2 * Scale * Scale);
            }
            internal double Area { get { return Math.Abs(Paths.Sum(RingArea)); } }
            internal double Perimeter
            { get { double sum = 0; foreach (var path in Paths) for (int i = 0; i < path.Count; i++) { var p = path[i]; var q = path[(i + 1) % path.Count]; double x = (double)p.X - q.X, y = (double)p.Y - q.Y; sum += Math.Sqrt(x * x + y * y) / Scale; } return sum; } }
            internal static PlanarRegion FromRings(IEnumerable<IEnumerable<double[]>> rings)
            {
                var paths = new List<List<CP>>();
                foreach (var ring in rings)
                {
                    var path = new List<CP>();
                    foreach (var p in ring) { var next = new CP(Coordinate(p[0]), Coordinate(p[1])); if (path.Count == 0 || path.Last() != next) path.Add(next); }
                    if (path.Count > 1 && path[0] == path.Last()) path.RemoveAt(path.Count - 1);
                    if (path.Count < 3 || RingArea(path) == 0) throw new InvalidOperationException("Контур вырожден, самопересекается или меньше точности 0,001 мм. Он не был молча удалён.");
                    paths.Add(path);
                }
                return Execute(paths, new List<List<CP>>(), CT.ctUnion, PF.pftEvenOdd);
            }
            private static PlanarRegion Execute(List<List<CP>> a, List<List<CP>> b, CT operation, PF fill)
            {
                if (a.Count == 0 && b.Count == 0) return Empty;
                var clipper = new PC { StrictlySimple = true, PreserveCollinear = false };
                clipper.AddPaths(a, PT.ptSubject, true); clipper.AddPaths(b, PT.ptClip, true);
                var result = new List<List<CP>>();
                if (!clipper.Execute(operation, result, fill, fill)) throw new InvalidOperationException("Не удалось выполнить плоскую операцию над замкнутыми контурами.");
                return new PlanarRegion(result);
            }
            internal static PlanarRegion Combine(PlanarRegion a, PlanarRegion b, CT operation)
            {
                a = a ?? Empty; b = b ?? Empty;
                if (b.IsEmpty) return operation == CT.ctIntersection ? Empty : a;
                if (a.IsEmpty) return operation == CT.ctUnion || operation == CT.ctXor ? b : Empty;
                // Normalized profiles have positive outer contours and negative holes.
                return Execute(a.Paths, b.Paths, operation, PF.pftNonZero);
            }
            internal IEnumerable<List<List<CP>>> Components()
            {
                var clipper = new PC { StrictlySimple = true }; clipper.AddPaths(Paths, PT.ptSubject, true);
                var tree = new TepClipper.PolyTree();
                if (!IsEmpty && !clipper.Execute(CT.ctUnion, tree, PF.pftNonZero, PF.pftNonZero)) throw new InvalidOperationException("Не удалось выделить компоненты контура.");
                for (var node = tree.GetFirst(); node != null; node = node.GetNext())
                    if (!node.IsHole && !node.IsOpen && node.Contour.Count > 0)
                    { var component = new List<List<CP>> { node.Contour }; component.AddRange(node.Childs.Where(n => n.IsHole).Select(n => n.Contour)); yield return component; }
            }
        }
        internal sealed class PlanarLayer
        {
            internal double Bottom, Top;
            internal PlanarRegion Region;
            internal double Volume { get { return Region.Area * (Top - Bottom); } }
        }
        internal sealed class LayeredBody
        {
            internal readonly List<PlanarLayer> Layers = new List<PlanarLayer>();
            internal bool Curved;
            internal double Volume { get { return Layers.Sum(l => l.Volume); } }
            internal PlanarRegion At(double z)
            { return Layers.FirstOrDefault(l => l.Bottom <= z && l.Top > z)?.Region ?? PlanarRegion.Empty; }
            internal PlanarRegion Projection()
            { var result = PlanarRegion.Empty; foreach (var l in Layers) result = PlanarRegion.Combine(result, l.Region, CT.ctUnion); return result; }
            internal static LayeredBody Combine(LayeredBody a, LayeredBody b, CT operation)
            {
                var levels = a.Layers.Concat(b.Layers).SelectMany(l => new[] { l.Bottom, l.Top }).Distinct().OrderBy(z => z).ToList();
                var result = new LayeredBody { Curved = a.Curved || b.Curved };
                for (int i = 0; i + 1 < levels.Count; i++)
                {
                    double z = levels[i] + (levels[i + 1] - levels[i]) / 2;
                    var area = PlanarRegion.Combine(a.At(z), b.At(z), operation);
                    if (!area.IsEmpty) result.Layers.Add(new PlanarLayer { Bottom = levels[i], Top = levels[i + 1], Region = area });
                }
                return result;
            }
        }

        public sealed class Engine
        {
            private readonly UIApplication app;
            private readonly UIDocument ui;
            private readonly Document doc;
            private readonly HashSet<long> selection;
            private static readonly Guid StorageId = new Guid("DAD13984-7185-4C89-9426-3043F1919CD2");
            private const string Owner = "KPLN.TEP.Spec20260826";
            public Settings Config { get; private set; } = new Settings();
            public List<Source> Sources { get; private set; } = new List<Source>();
            public List<string> Parameters { get; private set; } = new List<string>();
            public List<string> Categories { get; private set; } = new List<string>();
            public List<string> Phases { get; private set; } = new List<string>();
            public List<string> AreaSchemes { get; private set; } = new List<string>();
            public List<Choice> Templates { get; private set; } = new List<Choice>();
            public Run Last { get; private set; }
            public Detail RequestedDetail { get; set; }
            private Run current;
            private double zero, ground;
            private readonly Dictionary<string, Solid> shapes = new Dictionary<string, Solid>();
            private readonly HashSet<string> notices = new HashSet<string>();
            private readonly Dictionary<string, VolumeSet> volumeSets = new Dictionary<string, VolumeSet>(StringComparer.Ordinal);
            private readonly Dictionary<Element, List<Solid>> rawVolumeSolids = new Dictionary<Element, List<Solid>>();
            private readonly Dictionary<string, List<Solid>> worldVolumeSolids = new Dictionary<string, List<Solid>>(StringComparer.Ordinal);
            private readonly Dictionary<Tuple<Document, bool>, SpatialElementGeometryCalculator> spatialCalculators = new Dictionary<Tuple<Document, bool>, SpatialElementGeometryCalculator>();
            private readonly Dictionary<Tuple<Element, bool>, Solid> localSpatialVolumes = new Dictionary<Tuple<Element, bool>, Solid>();
            private readonly Dictionary<string, Timing> timings = new Dictionary<string, Timing>();
            private readonly List<SlowOperation> slowOperations = new List<SlowOperation>();
            private int volumeCacheHits, spatialCacheHits;
            private int booleanTouchSkips, booleanSplitRecoveries, booleanIntersectionRecoveries, booleanNormalizedRecoveries;
            private sealed class VolumeRetryBudget
            {
                private int remaining = 64;
                internal void Spend()
                { if (remaining-- <= 0) throw new InvalidOperationException("Исчерпан лимит 64 резервных геометрических операций для одной пары тел."); }
            }
            private sealed class VolumeFragment
            {
                internal string RecordKey;
                internal Solid Shape, Above, Below;
                internal bool Split;
            }
            private sealed class VolumeSet
            {
                internal readonly List<VolumeFragment> Fragments = new List<VolumeFragment>();
                internal readonly List<Tuple<string, string, string>> Problems = new List<Tuple<string, string, string>>();
                internal bool Reconstructed;
            }
            private sealed class Timing { internal long Ticks, Calls; }
            private sealed class SlowOperation
            { internal string Stage, Metric, Source, Building, Element; internal double Seconds; }
            private T Measure<T>(string stage, Record record, string metric, Func<T> action)
            {
                long start = System.Diagnostics.Stopwatch.GetTimestamp();
                try { return action(); }
                finally
                {
                    long ticks = System.Diagnostics.Stopwatch.GetTimestamp() - start; Timing timing;
                    if (!timings.TryGetValue(stage, out timing)) timings[stage] = timing = new Timing();
                    timing.Ticks += ticks; timing.Calls++;
                    double seconds = (double)ticks / System.Diagnostics.Stopwatch.Frequency;
                    if (seconds >= .5)
                    {
                        slowOperations.Add(new SlowOperation
                        {
                            Stage = stage,
                            Metric = metric,
                            Source = record?.Source.Name,
                            Building = record?.Building,
                            Element = record == null ? null : IDHelper.ElIdValue(record.Element.Id).ToString(),
                            Seconds = seconds
                        });
                        if (slowOperations.Count > 20) slowOperations.Remove(slowOperations.OrderBy(x => x.Seconds).First());
                    }
                }
            }
            private void WriteTimings()
            {
                foreach (var pair in timings.OrderByDescending(p => p.Value.Ticks))
                    current.Issue("PERF_STAGE", "Информация", pair.Key + ": " + ((double)pair.Value.Ticks / System.Diagnostics.Stopwatch.Frequency).ToString("0.###") + " с; операций: " + pair.Value.Calls,
                        action: "Времена этапов включают вложенные операции и не суммируются. Используйте для сравнения повторных запусков на одной модели.");
                current.Issue("PERF_CACHE", "Информация", "Повторно использовано наборов объёмных фрагментов: " + volumeCacheHits + "; локальных объёмов помещений / пространств: " + spatialCacheHits,
                    action: "Кэш действует только внутри одного запуска; набор фрагментов повторно используется при совпадении входных объектов и их расчётных назначений.");
                current.Issue("PERF_BOOLEAN", "Информация", "Резервная обработка объёмов: без объёмного пересечения - " + booleanTouchSkips + "; разбиением на связные части - " + booleanSplitRecoveries + "; через пересечение A ∩ B - " + booleanIntersectionRecoveries + "; в локальных координатах - " + booleanNormalizedRecoveries,
                    action: "Резервные операции запускаются после отказа прямого вычитания. Необработанные фрагменты сохраняют ошибку VOLUME_BOOLEAN.");
                current.Issue("PERF_PLANAR", "Информация", "Вычитаний объёмов по высотным участкам без Boolean Revit: " + planarVolumeCuts);
                foreach (var slow in slowOperations.OrderByDescending(x => x.Seconds))
                    current.Issue("PERF_SLOW", "Информация", slow.Stage + ": " + slow.Seconds.ToString("0.###") + " с", slow.Metric ?? "", slow.Source ?? "", slow.Building ?? "", slow.Element ?? "",
                        "Одна из 20 наиболее длительных операций от 0,5 с. Перейдите к объекту для проверки сложности геометрии.");
            }
            private void DisposeSpatialCalculators()
            {
                foreach (var calculator in spatialCalculators.Values) calculator.Dispose();
                spatialCalculators.Clear();
            }
            private void ClearVolumeCaches()
            {
                DisposeSpatialCalculators(); localSpatialVolumes.Clear(); rawVolumeSolids.Clear(); worldVolumeSolids.Clear(); volumeSets.Clear();
            }
            private readonly HashSet<Correction> appliedCorrections = new HashSet<Correction>();
            private readonly List<Issue> startupIssues = new List<Issue>();
            private bool settingsLoadFailed;
            public Engine(UIApplication application)
            {
                app = application; ui = app.ActiveUIDocument; doc = ui.Document;
                selection = new HashSet<long>(ui.Selection.GetElementIds().Select(IDHelper.ElIdValue));
            }
            public bool IsInitialized { get; private set; }
            private Action<string> reportProgress;
            private bool? createViewsForRun;
            private bool ViewsEnabled { get { return createViewsForRun ?? Config.CreateViews; } }
            private string progressMessage;
            private void Progress(string message) { progressMessage = message; reportProgress?.Invoke(message); }
            public void Initialize(Action<string> progress)
            {
                if (IsInitialized) return;
                reportProgress = progress;
                try
                {
                    Progress("Загрузка модели: поиск источников...");
                    Sources.Add(new Source { Key = "host", Name = doc.Title, Document = doc, Transform = Transform.Identity });
                    Discover(doc, "host", Transform.Identity, null, new HashSet<Document> { doc });
                    var documents = new Dictionary<Document, List<Element>>();
                    foreach (var source in Sources.Where(x => x.Loaded))
                    {
                        Progress("Загрузка модели: " + source.Name + " - чтение элементов...");
                        try
                        {
                            List<Element> elements;
                            if (!documents.TryGetValue(source.Document, out elements))
                            {
                                var ownIds = new HashSet<ElementId>(new FilteredElementCollector(source.Document).WherePasses(new ExtensibleStorageFilter(StorageId)).ToElementIds());
                                elements = new List<Element>(); int scanned = 0;
                                foreach (var e in new FilteredElementCollector(source.Document).WhereElementIsNotElementType())
                                {
                                    if (++scanned % 200 == 0) Progress("Загрузка модели: " + source.Name + " - прочитано объектов: " + scanned);
                                    if (e.Category != null && !(e is View) && !(e is DataStorage) && !ownIds.Contains(e.Id)) elements.Add(e);
                                }
                                documents.Add(source.Document, elements);
                            }
                            source.Elements = elements;
                        }
                        catch (System.OperationCanceledException) { throw; }
                        catch (Exception ex) { source.LoadError = "Не удалось прочитать элементы источника: " + ex.Message; }
                    }
                    try { Config = Read<Settings>("settings") ?? new Settings(); Normalize(Config); }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        Config = new Settings(); settingsLoadFailed = true; startupIssues.Add(new Issue
                        {
                            Code = "SETTINGS_RECOVERY",
                            Severity = "Предупреждение",
                            Message = ex.Message,
                            Action = "Загружены начальные настройки. Повреждённая запись не перезаписывается автоматически; проверьте настройки и явно сохраните их либо импортируйте JSON."
                        });
                    }
                    RestoreSourceOptions(); RefreshCatalogs();
                    Progress("Загрузка сохранённого отчёта...");
                    try { Last = Read<Run>("run"); }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { startupIssues.Add(new Issue { Code = "RUN_RECOVERY", Severity = "Предупреждение", Message = ex.Message, Action = "Выполните новый расчёт." }); }
                    IsInitialized = true;
                }
                finally { reportProgress = null; }
            }
            private void Discover(Document parent, string path, Transform transform, ElementId root, HashSet<Document> chain)
            {
                foreach (RevitLinkInstance link in new FilteredElementCollector(parent).OfClass(typeof(RevitLinkInstance)))
                {
                    Progress("Загрузка модели: поиск связи " + link.Name);
                    try
                    {
                        // Nested overlay links are not displayed through their parent attachment.
                        var type = parent.GetElement(link.GetTypeId()) as RevitLinkType;
                        if (root != null && type != null && type.AttachmentType == AttachmentType.Overlay) continue;
                        var linked = link.GetLinkDocument(); var key = path + "/" + link.UniqueId;
                        var source = new Source
                        {
                            Key = key,
                            Name = path == "host" ? link.Name : path + " / " + link.Name,
                            Document = linked,
                            Transform = transform.Multiply(link.GetTotalTransform()),
                            RootLink = root ?? link.Id
                        };
                        Sources.Add(source);
                        if (linked != null && !chain.Contains(linked)) { var next = new HashSet<Document>(chain) { linked }; Discover(linked, key, source.Transform, source.RootLink, next); }
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Sources.Add(new Source { Key = path + "/" + link.UniqueId, Name = link.Name, RootLink = root ?? link.Id, LoadError = ex.Message, Transform = transform }); }
                }
            }
            private static void Normalize(Settings s)
            {
                if (s.Version != 2) throw new InvalidOperationException("Формат настроек не соответствует этой версии ТЭП.");
                s.Metrics = s.Metrics ?? Catalog(); var enabled = s.Metrics.ToDictionary(x => x.Key, x => x);
                foreach (var official in Catalog())
                {
                    Metric m; if (enabled.TryGetValue(official.Key, out m)) { m.Name = official.Name; m.Unit = official.Unit; m.Description = official.Description; }
                    else s.Metrics.Add(official);
                }
                s.Metrics.RemoveAll(m => !Enum.IsDefined(typeof(Indicator), m.Key));
                s.Contours = s.Contours ?? new ObservableCollection<ContourSketch>();
                foreach (var m in s.Metrics)
                {
                    m.ContourMode = m.ContourMode ?? "current"; m.WallSelectionMode = m.WallSelectionMode ?? "exterior";
                    m.WallBoundary = m.WallBoundary ?? "auto"; m.WallCutHeight = m.WallCutHeight ?? "1";
                    m.ContourExclusions = m.ContourExclusions ?? "normative"; m.VolumeMaskHeight = m.VolumeMaskHeight ?? "actual";
                    m.SelectedWalls = m.SelectedWalls ?? new List<string>();
                    foreach (var pair in new[] { Tuple.Create("contour", m.ContourMode), Tuple.Create("wall-selection", m.WallSelectionMode), Tuple.Create("wall-boundary", m.WallBoundary), Tuple.Create("contour-exclusions", m.ContourExclusions), Tuple.Create("mask-height", m.VolumeMaskHeight) })
                        if (!Choices(pair.Item1).Any(x => x.Key == pair.Item2)) throw new InvalidOperationException("Неизвестная настройка контура: " + pair.Item1);
                    if (m.ContourMode != "current" && !SupportsContours(m.Key)) throw new InvalidOperationException("Для этого показателя нужен расчёт по классифицированным помещениям: " + m.Name);
                }
                foreach (var c in s.Contours)
                    if (!s.Metrics.Any(m => m.Key == c.Metric) || !Choices("sketch-kind").Any(x => x.Key == c.Kind)) throw new InvalidOperationException("Некорректная привязка ручного контура.");
                s.Parameters = s.Parameters ?? Settings.DefaultParameters();
                foreach (var p in Settings.DefaultParameters()) if (!s.Parameters.Any(x => x.Key == p.Key)) s.Parameters.Add(p);
                s.Rules = s.Rules ?? new ObservableCollection<Rule>(); s.Buildings = s.Buildings ?? new ObservableCollection<BuildingMap>();
                s.Levels = s.Levels ?? new ObservableCollection<LevelSetting>(); s.Corrections = s.Corrections ?? new ObservableCollection<Correction>();
                s.Sources = s.Sources ?? new List<SourceOption>(); s.Decimals = Math.Max(0, Math.Min(6, s.Decimals));
                var defaults = new Settings();
                foreach (var name in new[] { "IncludedColor", "ExcludedColor", "ManualColor", "PublicColor", "SummerColor", "TechnicalColor", "ErrorColor" })
                { var property = typeof(Settings).GetProperty(name); if (string.IsNullOrWhiteSpace((string)property.GetValue(s))) property.SetValue(s, property.GetValue(defaults)); }
                foreach (var name in new[] { "method", "profile", "group", "geometry", "graphics", "count" })
                {
                    var value = name == "method" ? s.Method : name == "profile" ? s.Profile : name == "group" ? s.Grouping : name == "geometry" ? s.Geometry : name == "graphics" ? s.Graphics : s.ApartmentMode;
                    if (!Choices(name).Any(x => x.Key == value)) throw new InvalidOperationException("Неизвестный вариант настройки: " + name + " = " + value);
                }
                // Earlier settings used a combo-box entry to disable graphics.
                if (s.Graphics == "none") { s.CreateViews = false; s.Graphics = "replace"; }
                foreach (var rule in s.Rules.Where(x => x.Enabled))
                {
                    if (rule.Metric != "all" && !s.Metrics.Any(m => m.Key == rule.Metric)) throw new InvalidOperationException("Неизвестный показатель правила: " + rule.Metric);
                    if (!Choices("operation").Any(x => x.Key == rule.Operation) || !Choices("action").Any(x => x.Key == rule.Action) || !Roles().Any(x => x.Key == rule.Role) || !Choices("part").Any(x => x.Key == rule.Part)) throw new InvalidOperationException("Неизвестное условие / действие / назначение в правиле.");
                    if (rule.Coefficient.HasValue && (double.IsNaN(rule.Coefficient.Value) || double.IsInfinity(rule.Coefficient.Value) || rule.Coefficient < 0 || rule.Coefficient > 1)) throw new InvalidOperationException("Коэффициент правила должен быть конечным числом от 0 до 1.");
                }
                foreach (var source in s.Sources) if (!Choices("source").Any(x => x.Key == source.Mode)) throw new InvalidOperationException("Неизвестный режим участия источника: " + source.Mode);
            }
            private void RestoreSourceOptions()
            {
                foreach (var s in Sources) { var saved = Config.Sources.FirstOrDefault(x => x.Key == s.Key); if (saved != null) { s.Mode = saved.Mode; s.Profile = saved.Profile; } }
            }
            private void SnapshotSources() { Config.Sources = Sources.Select(s => new SourceOption { Key = s.Key, Mode = s.Mode, Profile = s.Profile }).ToList(); }
            public void ImportSettings(string path)
            { var value = Deserialize<Settings>(File.ReadAllText(path, Encoding.UTF8)); Normalize(value); Config = value; settingsLoadFailed = false; RestoreSourceOptions(); RefreshCatalogs(); }
            public void ExportSettings(string path) { SnapshotSources(); File.WriteAllText(path, Serialize(Config), new UTF8Encoding(true)); }
            public void SaveSettings() { SnapshotSources(); Normalize(Config); Write("settings", Config); settingsLoadFailed = false; }
            public string StorageDescription { get { return "В текущем RVT: служебный элемент DataStorage «KPLN.TEP.Spec20260826/settings», Extensible Storage, схема " + StorageId + ", поле Payload. Это не параметр проекта. Для записи на диск сохраните RVT. Последний расчёт хранится аналогично в /run."; } }
            private bool catalogsReady;
            private void RefreshCatalogs()
            {
                var scannedDocuments = new HashSet<Document>();
                var names = new HashSet<string>(Parameters, StringComparer.OrdinalIgnoreCase) { "@Name", "@Category", "@Type", "@Family", "@Workset" };
                foreach (var s in Sources.Where(x => x.Loaded && x.LoadError == null))
                {
                    if (!catalogsReady && scannedDocuments.Add(s.Document))
                    {
                        int scanned = 0;
                        foreach (var e in s.Elements)
                        {
                            if (++scanned % 100 == 0) Progress("Загрузка параметров: " + s.Name + " - " + scanned + " / " + s.Elements.Count);
                            try { foreach (Parameter p in e.Parameters) if (p.Definition != null) names.Add(p.Definition.Name); }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { startupIssues.Add(new Issue { Code = "PARAM_SCAN", Severity = "Ошибка", Source = s.Name, Element = IDHelper.ElIdValue(e.Id).ToString(), Message = ex.Message, Action = "Проверьте повреждённый объект модели." }); }
                        }
                        foreach (ElementType e in new FilteredElementCollector(s.Document).WhereElementIsElementType())
                        {
                            if (++scanned % 100 == 0) Progress("Загрузка параметров типов: " + s.Name + " - " + scanned);
                            foreach (Parameter p in e.Parameters) if (p.Definition != null) names.Add(p.Definition.Name);
                        }
                    }
                    var occupied = new HashSet<ElementId>(s.Elements.OfType<SpatialElement>().Select(e => e.LevelId));
                    foreach (Level level in new FilteredElementCollector(s.Document).OfClass(typeof(Level)))
                    {
                        string key = s.Key + "/" + level.UniqueId;
                        var existing = Config.Levels.FirstOrDefault(x => x.Key == key && string.IsNullOrWhiteSpace(x.Building) && string.IsNullOrWhiteSpace(x.Section));
                        if (existing == null)
                        {
                            existing = new LevelSetting { Key = key }; Config.Levels.Add(existing);
                            var story = level.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY);
                            existing.Include = (story != null && story.AsInteger() == 1) || occupied.Contains(level.Id);
                        }
                        foreach (var setting in Config.Levels.Where(x => x.Key == key))
                        { setting.Source = s.Name; setting.Name = level.Name; setting.Elevation = s.Transform.OfPoint(new XYZ(0, 0, level.Elevation)).Z; }
                    }
                }
                catalogsReady = true;
                Parameters = names.OrderBy(x => x).ToList(); Categories = Sources.Where(s => s.Loaded).GroupBy(s => s.Document).Select(g => g.First()).SelectMany(s => s.Elements).Select(e => e.Category.Name).Distinct().OrderBy(x => x).ToList();
                Phases = Sources.Where(s => s.Loaded).GroupBy(s => s.Document).Select(g => g.First()).SelectMany(s => s.Document.Phases.Cast<Phase>()).Select(p => p.Name).Distinct().OrderBy(x => x).ToList();
                AreaSchemes = Sources.Where(s => s.Loaded).GroupBy(s => s.Document).Select(g => g.First()).SelectMany(s => new FilteredElementCollector(s.Document).OfClass(typeof(AreaScheme)).Cast<AreaScheme>()).Select(a => a.Name).Distinct().OrderBy(x => x).ToList();
                Templates = new List<Choice> { new Choice("", "Без шаблона", "Сохраняет цвета и оформление, созданные расчётом.") };
                Templates.AddRange(new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Where(v => v.IsTemplate).Select(v => new Choice(v.UniqueId, v.Name, "Применяет шаблон к служебному виду. Управляемые шаблоном категории и фильтры могут изменить видимость расчётной графики.")));
            }
            public List<Issue> CheckParameters(Action<string> progress = null)
            {
                parameterCache.Clear(); reportProgress = progress;
                try { return CheckParametersCore(); } finally { reportProgress = null; parameterCache.Clear(); }
            }
            private List<Issue> CheckParametersCore()
            {
                var r = new Run();
                foreach (var s in Sources.Where(x => x.Mode != "exclude"))
                {
                    if (!s.Loaded || s.LoadError != null) { r.Issue("SOURCE_UNLOADED", "Ошибка", s.LoadError ?? "Связь не загружена.", source: s.Name, action: "Загрузите связь или явно исключите её."); continue; }
                    foreach (var p in Config.Parameters.Where(x => !string.IsNullOrWhiteSpace(x.Name)))
                    {
                        int present = 0, filled = 0, ambiguous = 0;
                        int scanned = 0; foreach (var e in s.Elements) { if (++scanned % 100 == 0) Progress("Проверка параметра «" + p.Name + "»: " + s.Name + " - " + scanned + " / " + s.Elements.Count); try { var value = Value(e, p.Name); if (value != null) { present++; if (value.Trim().Length > 0) filled++; } } catch { ambiguous++; } }
                        r.Issue(present == 0 ? "PARAM_MISSING" : "PARAM_COVERAGE", present == 0 || ambiguous > 0 ? "Ошибка" : "Информация",
                            p.Title + " («" + p.Name + "»): есть у " + present + ", заполнен у " + filled + ", неоднозначен у " + ambiguous + " из " + s.Elements.Count + " объектов.", source: s.Name,
                            action: "Проверьте применимость параметра к нужным категориям. Одинаковые имена параметров нужно различить в модели.");
                    }
                    foreach (var rule in Config.Rules.Where(x => x.Enabled))
                    {
                        var candidates = s.Elements.Where(e => string.IsNullOrWhiteSpace(rule.Category) || e.Category.Name == rule.Category).ToList();
                        int hits = 0, missing = 0, checkedCount = 0; foreach (var e in candidates) { if (++checkedCount % 100 == 0) Progress("Проверка правил: " + s.Name + " - " + checkedCount + " / " + candidates.Count); try { if (Value(e, rule.Parameter) == null) missing++; else if (Matches(e, rule)) hits++; } catch { missing++; } }
                        r.Issue("RULE_CHECK", hits == 0 ? "Предупреждение" : "Информация", "Правило «" + rule.Parameter + " " + rule.Operation + " " + rule.Value + "»: совпало " + hits + ", параметр отсутствует или неоднозначен у " + missing + ".", rule.Metric, s.Name);
                    }
                }
                return r.Issues;
            }
            private static bool Eq(string a, string b) { return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase); }
            private readonly Dictionary<Element, Dictionary<string, IList<Parameter>>> parameterCache = new Dictionary<Element, Dictionary<string, IList<Parameter>>>();
            private readonly Dictionary<Document, Phase> phaseCache = new Dictionary<Document, Phase>();
            private IList<Parameter> NamedParameters(Element e, string name)
            {
                Dictionary<string, IList<Parameter>> names; if (!parameterCache.TryGetValue(e, out names)) parameterCache[e] = names = new Dictionary<string, IList<Parameter>>(StringComparer.Ordinal);
                IList<Parameter> result; if (!names.TryGetValue(name, out result)) names[name] = result = e.GetParameters(name); return result;
            }
            private Parameter Parameter(Element e, string name)
            {
                if (string.IsNullOrWhiteSpace(name) || name.StartsWith("@")) return null;
                var found = NamedParameters(e, name);
                if (found.Count == 0) { var type = e.Document.GetElement(e.GetTypeId()); if (type != null) found = NamedParameters(type, name); }
                if (found.Count > 1) throw new InvalidOperationException("Несколько параметров с именем «" + name + "».");
                return found.FirstOrDefault();
            }
            private string Value(Element e, string name)
            {
                if (string.IsNullOrWhiteSpace(name)) return null;
                if (name == "@Name") return e.Name;
                if (name == "@Category") return e.Category == null ? null : e.Category.Name;
                if (name == "@Type") return e.Document.GetElement(e.GetTypeId())?.Name;
                if (name == "@Family") return (e as FamilyInstance)?.Symbol.FamilyName;
                if (name == "@Workset") return e.Document.GetWorksetTable().GetWorkset(e.WorksetId)?.Name;
                var p = Parameter(e, name); if (p == null) return null; if (!p.HasValue) return "";
                if (p.StorageType == StorageType.String) return p.AsString() ?? "";
                if (p.StorageType == StorageType.Integer) return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                if (p.StorageType == StorageType.Double) return p.AsValueString() ?? p.AsDouble().ToString("R", CultureInfo.InvariantCulture);
                return p.AsValueString() ?? IDHelper.ElIdValue(p.AsElementId()).ToString();
            }
            private string Mapped(Element e, string key) { return Value(e, Config.Parameter(key)); }
            private double? Number(Element e, string key, bool length = false)
            {
                string name = Config.Parameter(key); var p = Parameter(e, name); if (p == null || !p.HasValue) return null;
                if (p.StorageType == StorageType.Double)
                {
                    if (length && !IDHelper.IsLength(p)) throw new InvalidOperationException("Параметр «" + name + "» должен иметь тип Длина, либо текст в метрах.");
                    if (!length && !IDHelper.IsNumber(p) && !IDHelper.IsAngle(p)) throw new InvalidOperationException("Параметр «" + name + "» должен быть безразмерным числом или углом.");
                    return p.AsDouble() * (length ? .3048 : IDHelper.IsAngle(p) ? 180 / Math.PI : 1);
                }
                if (length && p.StorageType != StorageType.String) throw new InvalidOperationException("Параметр «" + name + "» должен иметь тип Длина либо содержать текстовое число в метрах.");
                double n; return TryNumber(Value(e, name), out n) ? (double?)n : null;
            }
            public static bool TryNumber(string value, out double result)
            { return double.TryParse((value ?? "").Trim().Replace(",", "."), NumberStyles.Float, CultureInfo.InvariantCulture, out result) && !double.IsNaN(result) && !double.IsInfinity(result); }
            private static double RequiredNumber(string value, string title)
            { double n; if (!TryNumber(value, out n)) throw new InvalidOperationException("Укажите число: " + title + "."); return n; }
            private bool Flag(Element e, string key)
            { var v = Mapped(e, key); if (string.IsNullOrWhiteSpace(v)) return false; if (Eq(v, "1") || Eq(v, "true") || Eq(v, "да")) return true; if (Eq(v, "0") || Eq(v, "false") || Eq(v, "нет")) return false; throw new InvalidOperationException("Признак «" + Config.Parameter(key) + "» должен содержать да/нет или 1/0."); }
            private bool Matches(Element e, Rule rule)
            {
                if (!rule.Enabled || (!string.IsNullOrWhiteSpace(rule.Category) && !Eq(rule.Category, e.Category?.Name))) return false;
                string v = Value(e, rule.Parameter); if (v == null) return false;
                if (rule.Operation == "empty") return v.Trim().Length == 0; if (rule.Operation == "filled") return v.Trim().Length > 0;
                if (rule.Operation == "equals") return Eq(v, rule.Value);
                if (rule.Operation == "contains") return !string.IsNullOrEmpty(rule.Value) && v.IndexOf(rule.Value, StringComparison.OrdinalIgnoreCase) >= 0;
                throw new InvalidOperationException("Неизвестное условие правила: " + rule.Operation);
            }
            private static string RoleKey(string input)
            { return RoleLabels.FirstOrDefault(x => Eq(x.Key, input) || Eq(x.Value, input)).Key; }
            private void Notice(string code, string severity, string message, Record r = null, string metric = "", string action = "Уточните классификацию или исходную геометрию.")
            {
                string key = code + "|" + metric + "|" + (r == null ? "" : r.Key) + "|" + message; if (!notices.Add(key)) return;
                if (r?.Source.Mode == "reference" && severity == "Ошибка") severity = "Предупреждение";
                current.Issue(code, severity, message, metric, r?.Source.Name ?? "", r?.Building ?? "", r == null ? "" : IDHelper.ElIdValue(r.Element.Id).ToString(), action);
            }
            private bool PhaseAccepted(Source source, Element e)
            {
                if (e.DesignOption != null && !e.DesignOption.IsPrimary) return false;
                Phase phase; if (!phaseCache.TryGetValue(source.Document, out phase))
                { var phases = source.Document.Phases.Cast<Phase>().ToList(); phase = string.IsNullOrWhiteSpace(Config.Phase) ? phases.LastOrDefault() : phases.FirstOrDefault(p => Eq(p.Name, Config.Phase)); phaseCache[source.Document] = phase; }
                if (phase == null) throw new InvalidOperationException("В источнике отсутствует выбранная стадия «" + Config.Phase + "».");
                if (e is SpatialElement)
                { var parameter = e.get_Parameter(BuiltInParameter.ROOM_PHASE_ID); if (parameter != null && parameter.StorageType == StorageType.ElementId && parameter.AsElementId() != ElementId.InvalidElementId) return parameter.AsElementId() == phase.Id; }
                if (!e.HasPhases()) return true;
                var status = e.GetPhaseStatus(phase.Id); return status == ElementOnPhaseStatus.New || status == ElementOnPhaseStatus.Existing;
            }
            private Record MakeRecord(Source source, Element e)
            {
                var level = e.Document.GetElement(e.LevelId) as Level;
                if (level == null && e is SpatialElement) level = (e as SpatialElement).Level;
                string levelName = Mapped(e, "level");
                if (!string.IsNullOrWhiteSpace(levelName))
                {
                    var parameter = Parameter(e, Config.Parameter("level"));
                    if (parameter != null && parameter.StorageType == StorageType.ElementId) level = e.Document.GetElement(parameter.AsElementId()) as Level;
                    else level = new FilteredElementCollector(e.Document).OfClass(typeof(Level)).Cast<Level>().SingleOrDefault(l => Eq(l.Name, levelName));
                    if (level == null) throw new InvalidOperationException("Параметр расчётного уровня не соответствует уровню модели: " + levelName);
                }
                var setting = level == null ? null : Config.Levels.FirstOrDefault(x => x.Key == source.Key + "/" + level.UniqueId);
                var r = new Record { Source = source, Element = e, Level = setting, Section = Mapped(e, "section") ?? "", Apartment = Mapped(e, "apartment") ?? "", Vertical = Mapped(e, "vertical") ?? "", Role = "unknown", Part = "auto" };
                string value = Config.Grouping == "links" ? source.Name : Config.Grouping == "worksets" ? Value(e, "@Workset") :
                    Config.Grouping == "selection" ? Config.ManualBuilding : Mapped(e, "building");
                var maps = Config.Buildings.Where(x => (x.Source == "*" || Eq(x.Source, source.Key) || Eq(x.Source, source.Name)) && (x.MatchValue == "*" || Eq(x.MatchValue, value))).ToList();
                if (maps.Count > 1) throw new InvalidOperationException("Объект соответствует нескольким строкам назначения корпуса.");
                var map = maps.FirstOrDefault(); r.Building = map?.Building ?? value;
                if (string.IsNullOrWhiteSpace(r.Building)) throw new InvalidOperationException("Не определён корпус. Задайте параметр корпуса или таблицу соответствий.");
                r.BuildingIncluded = map?.Include ?? true;
                if (level != null)
                {
                    var overrides = Config.Levels.Where(l => l.Key == source.Key + "/" + level.UniqueId && (!string.IsNullOrWhiteSpace(l.Building) || !string.IsNullOrWhiteSpace(l.Section)) &&
                        (string.IsNullOrWhiteSpace(l.Building) || Eq(l.Building, r.Building)) && (string.IsNullOrWhiteSpace(l.Section) || Eq(l.Section, r.Section))).ToList();
                    if (overrides.Count > 1) throw new InvalidOperationException("Несколько настроек уровня соответствуют одному корпусу и секции.");
                    if (overrides.Count == 1) r.Level = overrides[0];
                }
                r.Profile = Config.Profile == "by-source" ? source.Profile : Config.Profile == "by-building" ? map?.Profile : Config.Profile;
                if (r.Profile == "high-mixed" && map != null && OneOf(map.Profile, "high-residential", "high-public")) r.Profile = map.Profile;
                string mappedProfile = Mapped(e, "profile"); if (!string.IsNullOrWhiteSpace(mappedProfile))
                { var p = Choices("profile").FirstOrDefault(x => Eq(x.Key, mappedProfile) || Eq(x.Label, mappedProfile)); if (p == null) throw new InvalidOperationException("Неизвестный профиль здания: " + mappedProfile); r.Profile = p.Key; }
                if (string.IsNullOrWhiteSpace(r.Profile) || r.Profile.StartsWith("by-")) throw new InvalidOperationException("Задайте профиль для корпуса / источника.");
                r.BuildingClass = map?.Class ?? (r.Profile.Contains("public") ? "nonresidential" : "residential");
                if (r.Profile == "high-mixed") throw new InvalidOperationException("Для многофункционального высотного здания назначьте элементам профиль функциональной части через параметр типа здания (high-residential / high-public). Требования СП 160 не подменяются одним общим правилом.");
                if (Config.Profile == "high-mixed") Notice("MIXED_PROFILE", "Предупреждение", "Для смешанного комплекса применён заданный профиль функциональной части. Требуется экспертная сверка состава частей по СП 160.", r);
                if (e is HostObject || e is Wall) r.Role = "structure";
                if (e.Category != null && IDHelper.ElIdValue(e.Category.Id) == (long)BuiltInCategory.OST_StructuralFoundation) r.Role = "structure";
                if (e.Category != null && IDHelper.ElIdValue(e.Category.Id) == (long)BuiltInCategory.OST_Parking) r.Role = "parking";
                if (e is Opening) r.Role = "opening";
                string function = Mapped(e, "function"); if (!string.IsNullOrWhiteSpace(function)) r.Role = RoleKey(function) ?? "unknown";
                string part = Mapped(e, "part"); if (!string.IsNullOrWhiteSpace(part)) r.Part = Choices("part").FirstOrDefault(x => Eq(x.Key, part) || Eq(x.Label, part))?.Key ?? "auto";
                ApplyRules(r, "all");
                if (r.Role == "unknown" && !string.IsNullOrWhiteSpace(r.Apartment))
                    Notice("APARTMENT_ROLE", "Ошибка", "ID квартиры найден, но назначение помещения не определено: нельзя отличить отапливаемую часть от летней.", r);
                return r;
            }
            private void ApplyRules(Record r, string metric)
            {
                var applicable = Config.Rules.Where(x => x.Metric == metric && Matches(r.Element, x)).ToList();
                var classifications = applicable.Where(x => x.Action == "classify").ToList();
                if (classifications.Select(x => x.Role + "|" + x.Part + "|" + x.Coefficient).Distinct().Count() > 1)
                    throw new InvalidOperationException("Конфликт классифицирующих правил для " + metric + ".");
                if (applicable.Any(x => x.Action == "include") && applicable.Any(x => x.Action == "exclude")) throw new InvalidOperationException("Одновременное включение и исключение правилами.");
                foreach (var rule in applicable)
                {
                    if (rule.Action == "classify") { r.Role = rule.Role; r.Part = rule.Part; r.Factor = rule.Coefficient; if (rule.Coefficient.HasValue) r.Manual = true; }
                    if (rule.Action == "include") { r.Override = true; r.Manual = true; }
                    if (rule.Action == "exclude") { r.Override = false; r.Manual = true; }
                }
                foreach (var c in Config.Corrections.Where(c => (c.Metric == metric || (metric == "all" && string.IsNullOrEmpty(c.Metric))) && c.Action != "delta" &&
                    (c.Source == r.Source.Key || Eq(c.Source, r.Source.Name)) && (c.Element == IDHelper.ElIdValue(r.Element.Id).ToString() || c.Element == r.Element.UniqueId)))
                {
                    if (string.IsNullOrWhiteSpace(c.Reason)) throw new InvalidOperationException("Для ручной корректировки необходима причина.");
                    if (c.Action == "include" || c.Action == "exclude") { c.Previous = r.Override?.ToString() ?? "По нормативу"; r.Override = c.Action == "include"; }
                    else if (c.Action == "classify") { c.Previous = r.Role; r.Role = RoleKey(c.Value) ?? throw new InvalidOperationException("Неизвестная роль корректировки."); }
                    else if (c.Action == "building") { c.Previous = r.Building; if (string.IsNullOrWhiteSpace(c.Value)) throw new InvalidOperationException("Пустой корпус в корректировке."); r.Building = c.Value; }
                    r.Manual = true;
                    appliedCorrections.Add(c);
                }
            }
            private double Datum(string mode, string value, string levelKey, string parameter, bool isGround)
            {
                if (mode == "manual") return RequiredNumber(value, isGround ? "отметка земли, м" : "отметка 0.000, м") / .3048;
                if (mode == "level")
                { var l = Config.Levels.FirstOrDefault(x => x.Key == levelKey); if (l == null) throw new InvalidOperationException("Не выбран уровень отсчёта."); return l.Elevation; }
                if (mode == "parameter")
                {
                    var p = Parameter(doc.ProjectInformation, parameter); if (p == null || !p.HasValue) throw new InvalidOperationException("Не заполнен параметр отметки в сведениях о проекте: " + parameter);
                    if (p.StorageType == StorageType.Double && IDHelper.IsLength(p)) return p.AsDouble();
                    return RequiredNumber(Value(doc.ProjectInformation, parameter), parameter) / .3048;
                }
                if (mode == "terrain" && isGround)
                {
                    var elevations = new List<double>();
                    foreach (var s in Sources.Where(x => x.Loaded && x.Mode == "include")) foreach (var e in s.Elements)
                        {
                            long category = IDHelper.ElIdValue(e.Category.Id);
                            // Toposolid is discovered by its runtime name to keep the same source compatible with Revit 2020.
                            if (category != (long)BuiltInCategory.OST_Topography && e.GetType().Name != "Toposolid") continue;
                            if (selection.Count > 0 && s.RootLink == null && !selection.Contains(IDHelper.ElIdValue(e.Id))) continue;
                            var top = e as TopographySurface;
                            if (top != null) elevations.AddRange(top.GetPoints().Select(p => s.Transform.OfPoint(p).Z));
                            else foreach (var solid in Solids(e)) foreach (Face face in solid.Faces)
                                        if (face.ComputeNormal(new UV(.5, .5)).Z > 0.1) elevations.AddRange(face.Triangulate().Vertices.Select(p => s.Transform.OfPoint(p).Z));
                        }
                    if (elevations.Count == 0) throw new InvalidOperationException("Не найдена поверхность земли. Выберите топографию до запуска или задайте отметку вручную.");
                    if (elevations.Max() - elevations.Min() > 0.01 / .3048)
                        throw new InvalidOperationException("Планировочная поверхность имеет уклон. Единую отметку земли нельзя определить по всему участку: задайте отметку вручную, а наземность этажей уточните по секциям.");
                    return elevations.Average();
                }
                throw new InvalidOperationException("Неизвестный способ определения отметки.");
            }
            public string PendingContourAction { get; set; }
            public string ContourMetricKey { get; set; } = "Gns";
            public bool ReopenContours { get; set; }
            private sealed class ContourSelectionFilter : ISelectionFilter
            {
                internal bool Walls;
                public bool AllowElement(Element e) { return Walls ? e is Wall : e is FilledRegion && (!Owned(e) || Kind(e) == "input-contour"); }
                public bool AllowReference(Reference reference, XYZ point) { return false; }
            }
            public void PerformContourAction()
            {
                string action = PendingContourAction; PendingContourAction = null; ReopenContours = true;
                try
                {
                    var metric = Config.Metrics.Single(m => m.Key == ContourMetricKey);
                    if (action == "walls")
                    {
                        var picked = ui.Selection.PickObjects(ObjectType.Element, new ContourSelectionFilter { Walls = true }, "Выберите стены активной модели и нажмите Готово");
                        metric.SelectedWalls = picked.Select(r => doc.GetElement(r).UniqueId).Distinct().ToList(); metric.WallSelectionMode = "selection"; metric.ContourMode = "walls"; return;
                    }
                    var regions = new List<FilledRegion>();
                    if (action.StartsWith("pick"))
                    {
                        var picked = ui.Selection.PickObjects(ObjectType.Element, new ContourSelectionFilter(), "Выберите цветовые области на поэтажных планах и нажмите Готово");
                        regions = picked.Select(r => doc.GetElement(r)).Cast<FilledRegion>().Distinct().ToList();
                        foreach (var region in regions)
                            if (!((doc.GetElement(region.OwnerViewId) as ViewPlan)?.GenLevel is Level)) throw new InvalidOperationException("Цветовая область должна находиться на плане с уровнем.");
                    }
                    else
                    {
                        var view = ui.ActiveView as ViewPlan;
                        if (view?.GenLevel == null || view.IsTemplate) throw new InvalidOperationException("Откройте поэтажный план нужного уровня перед запуском команды.");
                        var type = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).FirstOrDefault(e => !Owned(e))?.Id ?? ElementId.InvalidElementId;
                        if (type == ElementId.InvalidElementId) throw new InvalidOperationException("В проекте нет типа цветовой области. Создайте его средствами Revit.");
                        if (view.SketchPlane == null)
                            using (var t = new Transaction(doc, "ТЭП - плоскость ручного контура"))
                            { t.Start(); view.SketchPlane = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, view.GenLevel.Elevation))); t.Commit(); }
                        var points = new List<XYZ>();
                        while (true)
                        {
                            XYZ point;
                            try { point = ui.Selection.PickPoint(ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections, "Вершины контура; Esc после третьей точки замыкает контур"); }
                            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { if (points.Count < 3) return; break; }
                            point = new XYZ(point.X, point.Y, view.GenLevel.Elevation);
                            if (points.Count >= 3 && point.DistanceTo(points[0]) < doc.Application.ShortCurveTolerance) break;
                            if (points.Count == 0 || point.DistanceTo(points.Last()) > doc.Application.ShortCurveTolerance) points.Add(point);
                        }
                        var curves = points.Select((p, i) => (Curve)Line.CreateBound(p, points[(i + 1) % points.Count])).ToList();
                        var loops = new List<CurveLoop> { CurveLoop.Create(curves) }; Extrude(loops);
                        using (var t = new Transaction(doc, "ТЭП - ручной расчётный контур"))
                        { t.Start(); var region = FilledRegion.Create(doc, type, view.Id, loops); Tag(region, "input-contour", metric.Key); t.Commit(); regions.Add(region); }
                    }
                    string kind = action.EndsWith("exclude") ? "exclude" : "base";
                    if (kind == "base" && regions.Count > 0) metric.ContourMode = "manual";
                    foreach (var region in regions)
                    {
                        if (Config.Contours.Any(c => c.Metric == metric.Key && c.ElementUniqueId == region.UniqueId)) continue;
                        var level = ((ViewPlan)doc.GetElement(region.OwnerViewId)).GenLevel;
                        string profile = Config.Profile.StartsWith("by-") || Config.Profile == "high-mixed" ? "" : Config.Profile;
                        Config.Contours.Add(new ContourSketch
                        {
                            Metric = metric.Key,
                            ElementUniqueId = region.UniqueId,
                            ElementLabel = IDHelper.ElIdValue(region.Id).ToString(),
                            Kind = kind,
                            LevelKey = "host/" + level.UniqueId,
                            LevelName = level.Name,
                            Building = Config.Grouping == "links" ? doc.Title : Config.ManualBuilding,
                            Profile = profile,
                            BuildingClass = profile.Contains("public") ? "nonresidential" : "residential"
                        });
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
                catch (Exception ex) { TaskDialog.Show("Контуры ТЭП", ex.Message); }
            }
            private static bool ContourSourceLevel(string key, string source)
            { return key != null && key.StartsWith(source + "/", StringComparison.Ordinal) && key.IndexOf('/', source.Length + 1) < 0; }
            private LevelSetting ContourLevel(string key, string building, string section)
            {
                var settings = Config.Levels.Where(l => l.Key == key).ToList();
                var specific = settings.Where(l => (!string.IsNullOrWhiteSpace(l.Building) || !string.IsNullOrWhiteSpace(l.Section)) &&
                    (string.IsNullOrWhiteSpace(l.Building) || Eq(l.Building, building)) && (string.IsNullOrWhiteSpace(l.Section) || Eq(l.Section, section))).ToList();
                if (specific.Count > 1) throw new InvalidOperationException("Несколько настроек этажа соответствуют контуру: " + key);
                var result = specific.FirstOrDefault() ?? settings.FirstOrDefault(l => string.IsNullOrWhiteSpace(l.Building) && string.IsNullOrWhiteSpace(l.Section));
                if (result == null) throw new InvalidOperationException("Уровень контура отсутствует в настройках: " + key); return result;
            }
            private Record SketchRecord(ContourSketch sketch)
            {
                var region = doc.GetElement(sketch.ElementUniqueId) as FilledRegion;
                if (region == null) throw new InvalidOperationException("Цветовая область удалена или принадлежит другой модели: " + sketch.ElementLabel);
                var level = (doc.GetElement(region.OwnerViewId) as ViewPlan)?.GenLevel;
                if (level == null || sketch.LevelKey != "host/" + level.UniqueId) throw new InvalidOperationException("Изменился уровень ручного контура. Удалите привязку и выберите область заново.");
                if (string.IsNullOrWhiteSpace(sketch.Building) || !Choices("profile").Any(p => p.Key == sketch.Profile && !p.Key.StartsWith("by-") && p.Key != "high-mixed"))
                    throw new InvalidOperationException("Укажите корпус и однозначный профиль ручного контура: " + sketch.ElementLabel);
                if (!OneOf(sketch.BuildingClass, "residential", "nonresidential")) throw new InvalidOperationException("Укажите класс здания ручного контура.");
                return new Record
                {
                    Element = region,
                    Source = Sources.Single(s => s.Key == "host"),
                    Level = ContourLevel(sketch.LevelKey, sketch.Building, sketch.Section),
                    Building = sketch.Building,
                    Section = sketch.Section ?? "",
                    Profile = sketch.Profile,
                    BuildingClass = sketch.BuildingClass,
                    Role = "heated",
                    Part = "auto",
                    Manual = true
                };
            }
            private sealed class ContourInput
            {
                internal Record Record; internal Solid Plan; internal string Height; internal double Offset; internal bool Exclude;
            }
            private double ContourHeight(ContourInput input)
            {
                string value = string.IsNullOrWhiteSpace(input.Height) ? input.Record.Level.Height : input.Height;
                if (!string.IsNullOrWhiteSpace(value))
                { double height = RequiredNumber(value, "высота участка, м") / .3048 - (string.IsNullOrWhiteSpace(input.Height) ? input.Offset : 0); if (height <= 1e-6) throw new InvalidOperationException("Высота участка должна быть положительной."); return height; }
                var r = input.Record;
                var next = Config.Levels.Where(l => ContourSourceLevel(l.Key, r.Source.Key) && l.Elevation > r.Z + 1e-6)
                    .Select(l => l.Key).Distinct().Select(k => ContourLevel(k, r.Building, r.Section)).Where(l => l.Include && l.Kind != "exclude").OrderBy(l => l.Elevation).FirstOrDefault();
                if (next == null) throw new InvalidOperationException("Нет следующего расчётного уровня. Укажите высоту верхнего участка на вкладке «Этажи» или у ручного контура.");
                double result = next.Elevation - r.Z - input.Offset;
                if (result <= 1e-6) throw new InvalidOperationException("Смещение выреза находится выше следующего уровня. Задайте его высоту явно.");
                return result;
            }
            private static Solid ContourPrism(Solid plan, double bottom, double height)
            {
                // Disconnected islands are kept separate by the caller; holes remain in the face loops.
                return Extrude(BottomLoops(plan, bottom), height);
            }
            private static IEnumerable<Solid> PlanPieces(Solid plan)
            {
                foreach (var face in plan.Faces.Cast<Face>().OfType<PlanarFace>().Where(f => f.FaceNormal.Z < -.999999))
                    yield return Extrude(face.GetEdgesAsCurveLoops().Select(l => FlatLoop(l)).ToList());
            }
            private bool ContourWallSelected(Record r, Metric metric)
            {
                var wall = r.Element as Wall; if (wall == null) return false;
                if (metric.WallSelectionMode == "selection") return r.Source.Key == "host" && metric.SelectedWalls.Contains(wall.UniqueId);
                if (metric.WallSelectionMode == "exterior") return wall.WallType.Function == WallFunction.Exterior;
                if (string.IsNullOrWhiteSpace(metric.WallParameter) || string.IsNullOrWhiteSpace(metric.WallValue)) throw new InvalidOperationException("Задайте параметр и значение для отбора наружных стен.");
                string value = Value(wall, metric.WallParameter); return value != null && Eq(value, metric.WallValue);
            }
            private sealed class WallEdge
            { internal Record Record; internal Curve Curve; }
            private double WallOffset(WallEdge edge, Curve curve, Metric metric, Indicator indicator)
            {
                var wall = (Wall)edge.Record.Element;
                if (wall.WallType.Kind != WallKind.Basic || !(curve is Line)) throw new InvalidOperationException("Автоконтур поддерживает прямые вертикальные базовые стены. Для витражей, дуговых и сложных стен задайте ручную область.");
                BuiltInParameter sectionParameter;
                if (Enum.TryParse("WALL_CROSS_SECTION", out sectionParameter) && (wall.get_Parameter(sectionParameter)?.AsInteger() ?? 0) != 0)
                    throw new InvalidOperationException("Наклонная / переменная стена требует ручного контура.");
                string boundary = metric.WallBoundary;
                if (boundary == "auto") boundary = (int)indicator <= 4 ? (Config.Method == "new" ? "core" : "exterior") : GrossMetric(indicator) ? "interior" : "exterior";
                var side = boundary == "interior" ? ShellLayerType.Interior : ShellLayerType.Exterior;
                var right = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize().CrossProduct(XYZ.BasisZ);
                var offsets = new List<double>();
                foreach (var reference in HostObjectUtils.GetSideFaces(wall, side))
                {
                    var face = wall.GetGeometryObjectFromReference(reference) as PlanarFace;
                    if (face == null || Math.Abs(face.FaceNormal.Z) > 1e-8) continue;
                    var normal = edge.Record.Source.Transform.OfVector(face.FaceNormal).Normalize();
                    if (Math.Abs(normal.DotProduct(right)) < .999999) continue;
                    var origin = edge.Record.Source.Transform.OfPoint(face.Origin);
                    double distance = (origin - curve.GetEndPoint(0)).DotProduct(normal);
                    if (boundary == "core")
                    {
                        var compound = wall.WallType.GetCompoundStructure();
                        if (compound == null || compound.GetFirstCoreLayerIndex() < 0) throw new InvalidOperationException("У наружной стены не определено ядро.");
                        distance -= compound.GetLayers().Take(compound.GetFirstCoreLayerIndex()).Sum(l => l.Width);
                    }
                    offsets.Add(distance * normal.DotProduct(right));
                }
                if (offsets.Count == 0 || offsets.Max() - offsets.Min() > 1e-6) throw new InvalidOperationException("Не удалось однозначно определить боковую поверхность стены " + IDHelper.ElIdValue(wall.Id) + ". Используйте ручной контур.");
                return offsets[0];
            }
            private Solid WallContour(List<Record> records, Metric metric, Indicator indicator)
            {
                var remaining = records.Select(r => new WallEdge { Record = r, Curve = Flat(((LocationCurve)r.Element.Location).Curve.CreateTransformed(r.Source.Transform)) }).ToList();
                var loops = new List<CurveLoop>(); const double tolerance = 1e-5;
                // Degree two is required. Branches and gaps must be corrected rather than silently bridged.
                foreach (var edge in remaining)
                    for (int end = 0; end < 2; end++)
                    {
                        int count = remaining.Where(e => !ReferenceEquals(e, edge)).Sum(e => (e.Curve.GetEndPoint(0).DistanceTo(edge.Curve.GetEndPoint(end)) < tolerance ? 1 : 0) + (e.Curve.GetEndPoint(1).DistanceTo(edge.Curve.GetEndPoint(end)) < tolerance ? 1 : 0));
                        if (count != 1) throw new InvalidOperationException("Разрыв или ответвление наружных стен у ID " + IDHelper.ElIdValue(edge.Record.Element.Id) + ". Нужен замкнутый контур без ответвлений.");
                    }
                while (remaining.Count > 0)
                {
                    Progress("Сборка колец наружных стен: осталось " + remaining.Count);
                    var first = remaining[0]; remaining.RemoveAt(0);
                    var curves = new List<Curve> { first.Curve }; var offsets = new List<double> { WallOffset(first, first.Curve, metric, indicator) };
                    while (curves.Last().GetEndPoint(1).DistanceTo(curves[0].GetEndPoint(0)) >= tolerance)
                    {
                        var end = curves.Last().GetEndPoint(1);
                        var edge = remaining.Single(e => e.Curve.GetEndPoint(0).DistanceTo(end) < tolerance || e.Curve.GetEndPoint(1).DistanceTo(end) < tolerance);
                        var curve = edge.Curve.GetEndPoint(0).DistanceTo(end) < tolerance ? edge.Curve : edge.Curve.CreateReversed();
                        curves.Add(curve); offsets.Add(WallOffset(edge, curve, metric, indicator)); remaining.Remove(edge);
                    }
                    if (curves.Count < 3) throw new InvalidOperationException("Недостаточно стен для замкнутого контура.");
                    loops.Add(CurveLoop.CreateViaOffset(CurveLoop.Create(curves), offsets, XYZ.BasisZ));
                }
                return Extrude(loops);
            }
            private List<ContourInput> ContourInputs(List<Record> records, Metric metric, Indicator indicator)
            {
                var result = new List<ContourInput>();
                if (metric.ContourMode == "walls")
                {
                    if (metric.WallSelectionMode == "selection")
                        foreach (var id in metric.SelectedWalls.Where(id => !records.Any(r => r.Source.Key == "host" && r.Element.UniqueId == id)))
                            Notice("CONTOUR_WALL_MISSING", "Ошибка", "Выбранная стена отсутствует в доступном составе расчёта: " + id, metric: metric.Key);
                    var selected = records.Where(r => ContourWallSelected(r, metric)).ToList();
                    double cut = RequiredNumber(metric.WallCutHeight, "высота сечения стен, м") / .3048;
                    if (cut < 0) throw new InvalidOperationException("Высота сечения стен не может быть отрицательной.");
                    foreach (var group in selected.GroupBy(r => new { r.Source.Key, r.Building, r.Section, r.Profile, r.BuildingClass }))
                    {
                        var seed = group.First();
                        var levels = Config.Levels.Where(l => ContourSourceLevel(l.Key, seed.Source.Key)).Select(l => l.Key).Distinct()
                            .Select(k => ContourLevel(k, seed.Building, seed.Section)).Where(l => l.Include && l.Kind != "exclude").OrderBy(l => l.Elevation);
                        foreach (var level in levels)
                        {
                            Progress("Наружный контур: " + seed.Building + " / " + level.Name);
                            var r = seed.Copy(); r.Level = level; r.Override = null; r.Role = "heated"; r.Factor = null;
                            try
                            {
                                double z = level.Elevation + cut;
                                var onLevel = group.Where(w =>
                                {
                                    var box = w.Element.get_BoundingBox(null); if (box == null) return false;
                                    var a = w.Source.Transform.OfPoint(box.Transform.OfPoint(box.Min)); var b = w.Source.Transform.OfPoint(box.Transform.OfPoint(box.Max));
                                    return z >= Math.Min(a.Z, b.Z) - 1e-6 && z < Math.Max(a.Z, b.Z) - 1e-6;
                                }).ToList();
                                if (onLevel.Count == 0)
                                {
                                    if (records.Any(x => x.Source.Key == r.Source.Key && x.Building == r.Building && x.Section == r.Section && x.Level?.Key == level.Key && x.Element is SpatialElement))
                                        Notice("CONTOUR_WALL_LEVEL", "Ошибка", "На этаже есть помещения, но на высоте сечения нет выбранных стен.", r, metric.Key);
                                    continue;
                                }
                                var plan = WallContour(onLevel, metric, indicator);
                                foreach (var piece in PlanPieces(plan)) result.Add(new ContourInput { Record = r, Plan = piece });
                            }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { Notice("CONTOUR_WALLS", "Ошибка", level.Name + ": " + ex.Message, r, metric.Key); }
                        }
                    }
                }
                foreach (var sketch in Config.Contours.Where(c => c.Metric == metric.Key && (c.Kind == "exclude" || metric.ContourMode == "manual")))
                {
                    Record r = null;
                    try
                    {
                        r = SketchRecord(sketch); if (r.Source.Mode == "exclude" || !r.Level.Include || r.Level.Kind == "exclude") continue;
                        double offset = string.IsNullOrWhiteSpace(sketch.BottomOffset) ? 0 : RequiredNumber(sketch.BottomOffset, "смещение низа выреза, м") / .3048;
                        if (sketch.Kind == "base" && Math.Abs(offset) > 1e-8) throw new InvalidOperationException("Смещение низа допускается только у выреза. Основной контур начинается на отметке своего уровня.");
                        var shape = Extrude(((FilledRegion)r.Element).GetBoundaries().Select(l => FlatLoop(l)).ToList());
                        foreach (var piece in PlanPieces(shape)) result.Add(new ContourInput { Record = r, Plan = piece, Height = sketch.Height, Offset = offset, Exclude = sketch.Kind == "exclude" });
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("CONTOUR_SKETCH", "Ошибка", "Область " + sketch.ElementLabel + ": " + ex.Message, r, metric.Key); if (sketch.Kind == "exclude") throw new InvalidOperationException("Ручной вырез недоступен. Расчёт по этому показателю остановлен до восстановления выреза.", ex); }
                }
                if (!result.Any(i => !i.Exclude)) Notice("CONTOUR_EMPTY", "Ошибка", "Не найдено ни одного основного контура. Проверьте выбор стен / областей, уровни и источники.", metric: metric.Key);
                return result;
            }
            private bool ContourMaskRequired(Record r, Metric metric, Indicator indicator, List<Record> all)
            {
                if (r.Override.HasValue) return !r.Override.Value;
                string include = Mapped(r.Element, "include"); if (Eq(include, "0") || Eq(include, "нет") || Eq(include, "false")) return true;
                if (metric.ContourExclusions == "rules" || indicator == Indicator.Footprint) return false;
                if (OneOf(r.Role, "structure", "envelope", "footprint", "underground-footprint")) return false;
                if (r.Role == "unknown")
                { if (r.Element is SpatialElement) throw new InvalidOperationException("Назначение помещения / зоны не определено. Нельзя проверить нормативные вырезы внутри контура."); return false; }
                if (metric.Key.StartsWith("Volume")) return OneOf(r.Role, "balcony", "terrace", "canopy", "passage", "decoration", "ventilated-void", "soil-filled");
                if (!Eligible(r, indicator)) return r.Element is SpatialElement || OneOf(r.Role, "multilight", "stair-gap", "opening", "shaft", "engineering-shaft", "stove", "decoration");
                if (OneOf(r.Role, "multilight", "stair-gap", "opening", "shaft", "engineering-shaft")) return VerticalExclusion(r, indicator, all, ContourMaskPlan(r, indicator));
                return false;
            }
            private Solid ContourMaskPlan(Record r, Indicator indicator)
            {
                string key = r.Key + "/contour-mask"; Solid plan;
                if (!shapes.TryGetValue(key, out plan))
                {
                    plan = r.Element is SpatialElement ? SpatialPlan(r, "net", indicator.ToString()) : Plan(r, Indicator.Footprint);
                    if (plan == null || plan.Volume < 1e-9) throw new InvalidOperationException("Пустая геометрия исключаемого объекта.");
                    shapes[key] = plan;
                }
                return plan;
            }
            private bool ContourBaseIncluded(Record r, Indicator indicator)
            {
                if (r.Level == null || !r.Level.Include || r.Level.Kind == "exclude") return false;
                if (indicator == Indicator.Footprint || indicator.ToString().StartsWith("Volume")) return true;
                var probe = r.Copy(); probe.Role = r.Level.Kind == "attic" ? "attic" : r.Level.Kind == "void" ? "technical-void" : r.Level.Kind == "roof" ? "roof-used" : r.Level.Kind == "technical" ? "technical-room" : "heated";
                // A wall's inclusion parameter must not discard an entire floor envelope.
                var key = r.Level.Key.Substring(r.Source.Key.Length + 1); probe.Element = r.Source.Document.GetElement(key) ?? r.Element;
                probe.Override = null; return Eligible(probe, indicator);
            }
            private void CalculateContours(List<Record> records, Indicator indicator, Metric metric)
            {
                bool volume = metric.Key.StartsWith("Volume"), footprint = indicator == Indicator.Footprint;
                var inputs = ContourInputs(records, metric, indicator); var bases = inputs.Where(i => !i.Exclude).ToList();
                if (volume) Notice("CONTOUR_PRISMS", "Предупреждение", "Расчёт по вертикальным поэтажным призмам: контур постоянен по высоте участка. Наклонные кровли, наклонные стены и изменения сечения внутри этажа этим режимом не восстанавливаются. Проверьте высоты и 3D либо используйте текущий способ по геометрии.", metric: metric.Key);
                if (metric.ContourExclusions == "rules") Notice("CONTOUR_EXPLICIT_ONLY", "Предупреждение", "Автоматические исключения по назначению отключены. Применяются только явные правила, параметр включения и ручные вырезы.", metric: metric.Key);
                if (metric.ContourExclusions == "normative" && !records.Any(r => r.Element is SpatialElement))
                    Notice("CONTOUR_NO_ROOMS", "Предупреждение", "В составе расчёта нет помещений, пространств или зон для проверки исключений. Проверьте полноту вырезов вручную.", metric: metric.Key);
                var maskRecords = new List<Record>(); var badGroups = new HashSet<string>();
                Func<Record, string> groupKey = r => r.Building + "|" + r.Section + "|" + (r.Source.Mode == "reference" ? r.Source.Key : "include");
                foreach (var r in records.Where(r => r.Level != null && r.Level.Include && r.Level.Kind != "exclude"))
                {
                    if (!bases.Any(b => groupKey(b.Record) == groupKey(r) && (volume || footprint || Math.Abs(b.Record.Z - r.Z) < 1e-6))) continue;
                    try { if (ContourMaskRequired(r, metric, indicator, records)) maskRecords.Add(r); }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("CONTOUR_CLASSIFICATION", "Ошибка", ex.Message, r, metric.Key); badGroups.Add(groupKey(r)); }
                }
                var used = new Dictionary<string, PlanIndex>(); int count = 0;
                foreach (var input in bases.OrderBy(i => i.Record.Z).ThenBy(i => i.Record.Key, StringComparer.Ordinal))
                {
                    var r = input.Record; string key = groupKey(r);
                    Progress("Контур и вырезы: " + metric.Name + " - " + (++count) + " / " + bases.Count + "; " + r.Level.Name);
                    try
                    {
                        if (!ContourBaseIncluded(r, indicator) || badGroups.Contains(key)) continue;
                        if (footprint && r.Z > ground + 1e-6) continue;
                        double height = volume ? ContourHeight(input) : 1;
                        Solid shape = volume ? ContourPrism(input.Plan, r.Z, height) : input.Plan;
                        // Footprint uses the union of contours at/below ground; upper protrusions must be supplied as explicit ground outlines.
                        var maskIndex = new PlanIndex(volume); var exclusions = new List<Detail>();
                        Action<Record, Solid, string> addMask = (owner, mask, reason) =>
                        {
                            if (mask == null || mask.Volume < 1e-9) return;
                            maskIndex.Add(mask, owner);
                            var intersections = volume ? VolumeIntersections(shape, mask) : new List<Solid> { Intersect(shape, mask) };
                            foreach (var part in intersections)
                            {
                                var clipped = part; if (clipped == null || clipped.Volume < 1e-9) continue;
                                if (volume && indicator != Indicator.Volume) clipped = Half(clipped, zero, indicator == Indicator.VolumeAbove);
                                if (clipped == null || clipped.Volume < 1e-9) continue;
                                var row = Row(owner, indicator, clipped.Volume * (volume ? .028316846592 : .09290304), 1, true, reason, clipped, volume);
                                row.Level = r.Level.Name; row.Elevation = footprint ? ground : r.Z; exclusions.Add(row);
                            }
                        };
                        foreach (var mask in maskRecords.Where(m => groupKey(m) == key && (volume || footprint || Math.Abs(m.Z - r.Z) < 1e-6)))
                        {
                            Progress("Вырезы: " + r.Level.Name + "; ID " + IDHelper.ElIdValue(mask.Element.Id));
                            if (volume && metric.VolumeMaskHeight == "actual")
                            {
                                if (mask.Element is Area) throw new InvalidOperationException("Зона ID " + IDHelper.ElIdValue(mask.Element.Id) + " не имеет объёма. Выберите вырез на всю высоту участка или задайте ручной вырез.");
                                var solids = mask.Element is SpatialElement ? new List<Solid> { SpatialVolume(mask) } : VolumeSolids(mask, metric.Key);
                                if (solids.Count == 0) throw new InvalidOperationException("Нет объёмной геометрии выреза ID " + IDHelper.ElIdValue(mask.Element.Id));
                                foreach (var solid in solids) addMask(mask, solid, "Исключение по фактической геометрии объекта");
                            }
                            else
                            {
                                if (volume && Math.Abs(mask.Z - r.Z) > 1e-6) continue;
                                foreach (var plan in PlanPieces(ContourMaskPlan(mask, indicator)))
                                    addMask(mask, volume ? ContourPrism(plan, r.Z, height) : plan, volume ? "Исключение на всю высоту участка" : "Исключение по контуру помещения / элемента");
                            }
                        }
                        foreach (var mask in inputs.Where(m => m.Exclude && groupKey(m.Record) == key && (volume || footprint || Math.Abs(m.Record.Z - r.Z) < 1e-6)))
                        {
                            var solid = volume ? ContourPrism(mask.Plan, mask.Record.Z + mask.Offset, ContourHeight(mask)) : mask.Plan;
                            addMask(mask.Record, solid, "Ручной вырез цветовой областью");
                        }
                        string usedKey = key + (volume || footprint ? "" : "|" + r.Z.ToString("R", CultureInfo.InvariantCulture));
                        PlanIndex previous; if (!used.TryGetValue(usedKey, out previous)) used[usedKey] = previous = new PlanIndex(volume);
                        var fragments = volume ? RemoveNearbyVolumes(new List<Solid> { shape }, maskIndex, r, metric.Key, "Исключения внутри контура") : new List<Solid> { RemoveNearby(shape, maskIndex, r, metric.Key, "Вырезы контура") };
                        var final = new List<Solid>();
                        foreach (var fragment in fragments)
                        {
                            if (fragment == null || fragment.Volume < 1e-9) continue;
                            if (volume) final.AddRange(RemoveNearbyVolumes(new List<Solid> { fragment }, previous, r, metric.Key, "Устранение повторного учёта контуров"));
                            else { var unique = RemoveNearby(fragment, previous, r, metric.Key, "Пересечения контуров"); if (unique != null && unique.Volume > 1e-9) final.Add(unique); }
                        }
                        // Publish the whole input atomically, including cuts at zero, before indexing it as counted.
                        var rows = new List<Detail>();
                        foreach (var fragment in final)
                        {
                            var measured = volume && indicator != Indicator.Volume ? Half(fragment, zero, indicator == Indicator.VolumeAbove) : fragment;
                            if (measured == null || measured.Volume < 1e-9) continue;
                            var row = Row(r, indicator, measured.Volume * (volume ? .028316846592 : .09290304), 1, r.Source.Mode == "reference",
                                (metric.ContourMode == "walls" ? "Контур наружных стен" : "Ручной контур") + "; вырезы и пересечения вычтены" + (volume ? "; высота участка " + (height * .3048).ToString("0.###", CultureInfo.InvariantCulture) + " м" : ""), measured, volume);
                            row.Purpose = "envelope";
                            if (footprint) { row.Level = "План застройки"; row.Elevation = ground; }
                            rows.Add(row);
                        }
                        foreach (var fragment in final) previous.Add(fragment, r);
                        current.Details.AddRange(exclusions); current.Details.AddRange(rows);
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("CONTOUR_CALCULATION", "Ошибка", r.Level.Name + ": " + ex.Message, r, metric.Key); }
                }
                if (footprint && !bases.Any(b => b.Record.Source.Mode == "include" && b.Record.Z <= ground + 1e-6))
                    Notice("CONTOUR_GROUND_EMPTY", "Ошибка", "Нет основного контура на отметке земли или ниже. Проверьте отметку земли и уровни областей.", metric: metric.Key);
                if (footprint) Notice("CONTOUR_FOOTPRINT_EXTRAS", "Предупреждение", "В этом режиме площадь застройки определяется выбранными контурами: наземным сечением и проекциями подземных этажей. Крыльца, приямки, консоли и другие выступающие части добавьте в ручные основные области либо используйте текущий способ.", metric: metric.Key);
            }

            private List<Record> Collect()
            {
                var records = new List<Record>();
                foreach (var source in Sources.Where(s => s.Mode != "exclude"))
                {
                    if (!source.Loaded || source.LoadError != null) { current.Issue("SOURCE_UNLOADED", "Ошибка", source.LoadError ?? "Источник не загружен", source: source.Name, action: "Загрузите / восстановите связь либо явно исключите её из расчёта."); continue; }
                    if (Math.Abs(source.Transform.BasisZ.DotProduct(XYZ.BasisZ) - 1) > 1e-8)
                    { Notice("SOURCE_TILTED", "Ошибка", "Наклонённая связь не поддерживает поэтажную классификацию: " + source.Name); continue; }
                    var candidates = new List<Record>();
                    int scanned = 0;
                    foreach (var e in source.Elements)
                    {
                        if (++scanned % 100 == 0) Progress("Классификация: " + source.Name + " - " + scanned + " / " + source.Elements.Count);
                        if (!(e is SpatialElement) && !(e is HostObject) && !(e is FamilyInstance) && !(e is Opening) && !(e is DirectShape) && !(e is Ceiling)) continue;
                        if (Config.Grouping == "selection" && !selection.Contains(IDHelper.ElIdValue(source.RootLink ?? e.Id))) continue;
                        try
                        {
                            if (!PhaseAccepted(source, e)) continue;
                            bool parking = e.Category != null && IDHelper.ElIdValue(e.Category.Id) == (long)BuiltInCategory.OST_Parking;
                            bool foundation = e.Category != null && IDHelper.ElIdValue(e.Category.Id) == (long)BuiltInCategory.OST_StructuralFoundation;
                            bool configured = Config.Rules.Any(rule => rule.Enabled && (rule.Metric == "all" || Config.Metrics.Any(m => m.Enabled && m.Key == rule.Metric)) && Matches(e, rule));
                            bool hasFunction = !string.IsNullOrWhiteSpace(Mapped(e, "function"));
                            bool construction = Config.Metrics.Any(m => m.Enabled && (m.ContourMode == "walls" || m.Key.StartsWith("Volume") || OneOf(m.Key, "Footprint", "Storeys", "Floors")));
                            if (!(e is SpatialElement) && !(e is Opening) && !parking && !hasFunction && !configured && !(construction && (e is HostObject || foundation))) continue;
                            var spatial = e as SpatialElement;
                            if (spatial != null && SpatialArea(spatial) <= 1e-9) { current.Issue("SPATIAL_UNBOUNDED", "Ошибка", "Не размещённое или не замкнутое помещение / пространство / зона.", source: source.Name, element: IDHelper.ElIdValue(e.Id).ToString()); continue; }
                            var record = MakeRecord(source, e); if (record.BuildingIncluded) candidates.Add(record);
                        }
                        catch (System.OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            var affected = (e is HostObject) ? Config.Metrics.Where(m => m.Enabled && (m.ContourMode == "walls" || m.Key.StartsWith("Volume") || OneOf(m.Key, "Footprint", "Storeys", "Floors"))).Select(m => m.Key).ToList() : new List<string> { "" };
                            foreach (var metric in affected) current.Issue("CLASSIFICATION", source.Mode == "reference" ? "Предупреждение" : "Ошибка", ex.Message, metric, source.Name, element: IDHelper.ElIdValue(e.Id).ToString());
                        }
                    }
                    records.AddRange(candidates);
                }
                if (Config.Grouping == "selection" && selection.Count == 0) Notice("SELECTION_EMPTY", "Ошибка", "Ручной режим требует выбора объектов или экземпляров связей до запуска команды.");
                return records;
            }
            private List<Record> Representations(List<Record> records, Metric metric)
            {
                var selectedSpatial = new HashSet<Record>();
                foreach (var original in records.Where(r => r.Element is SpatialElement).GroupBy(r => r.Source.Key + "|" + r.Level?.Key))
                {
                    string scheme = string.IsNullOrWhiteSpace(metric.AreaScheme) ? Config.AreaScheme : metric.AreaScheme;
                    var group = original.Where(r => !(r.Element is Area) || string.IsNullOrWhiteSpace(scheme) || Eq(((Area)r.Element).AreaScheme.Name, scheme)).ToList();
                    string kind = string.IsNullOrWhiteSpace(metric.Geometry) ? Config.Geometry : metric.Geometry;
                    if (metric.ContourMode == "current" && (metric.Key.StartsWith("Volume") || metric.Key == "Footprint") || metric.ContourMode != "current" && kind == "auto")
                        kind = group.Any(r => r.Element is Room) ? "rooms" : group.Any(r => r.Element is Space) ? "spaces" : kind;
                    if (kind == "auto") kind = group.Any(r => r.Element is Area) ? "areas" : group.Any(r => r.Element is Room) ? "rooms" : "spaces";
                    var chosen = group.Where(r => kind == "areas" ? r.Element is Area : kind == "rooms" ? r.Element is Room : r.Element is Space).ToList();
                    if (kind == "areas" && chosen.Select(r => ((Area)r.Element).AreaScheme.Id).Distinct().Count() > 1)
                    { Notice("AREA_SCHEME", "Ошибка", "На уровне несколько схем зон. Выберите одну расчётную схему.", group.First(), metric.Key); continue; }
                    if (chosen.Count == 0 && group.Count > 0) Notice("REPRESENTATION_MISSING", "Ошибка", "На уровне отсутствует выбранный источник площадей: " + kind, group.First(), metric.Key);
                    foreach (var r in chosen) selectedSpatial.Add(r);
                    if (chosen.Count < original.Count()) Notice("REPRESENTATION", "Информация", "На уровне используется один источник площадей: " + kind + ". Параллельные Rooms/Spaces/Areas не суммируются.", original.First(), metric.Key);
                }
                return records.Where(r => !(r.Element is SpatialElement) || selectedSpatial.Contains(r)).ToList();
            }
            private static double SpatialArea(SpatialElement e)
            { return e is Room ? ((Room)e).Area : e is Space ? ((Space)e).Area : e is Area ? ((Area)e).Area : 0; }
            private static bool IsPublic(Record r) { return r.Profile == "public" || r.Profile == "high-public"; }
            private static bool IsHigh(Record r) { return r.Profile.StartsWith("high-"); }
            private static bool OneOf(string value, params string[] values) { return values.Contains(value); }
            private static bool GrossMetric(Indicator metric)
            { return OneOf(metric.ToString(), "Gross", "GrossAbove", "GrossBelow", "Np", "NpResidential", "NpNonresidential", "NnpEmbedded", "NnpSeparate"); }
            private bool Above(Record r, bool forStoreys = false)
            {
                if (r.Level == null) throw new InvalidOperationException("Не определён расчётный этаж объекта.");
                if (r.Level.Above == "above") return true; if (r.Level.Above == "below") return false;
                if (r.Level.Kind == "basement")
                {
                    if (forStoreys && IsPublic(r) && !IsHigh(r)) return true;
                    double top = RequiredNumber(r.Level.TopSlab, "отметка верха перекрытия цоколя «" + r.Level.Name + "», м") / .3048;
                    return top - ground >= 2 / .3048 - 1e-8;
                }
                return r.Z >= ground - 1e-6;
            }
            private bool CountLevel(Record r, bool aboveOnly)
            {
                if (r.Level == null || !r.Level.Include || r.Level.Kind == "exclude") return false;
                if (r.Level.Kind == "attic") return false;
                if (r.Level.Kind == "roof" || OneOf(r.Role, "roof-exit", "roof-vent"))
                {
                    if (IsHigh(r))
                    {
                        double area = RequiredNumber(r.Level.RoofArea, "суммарная площадь верхней надстройки, м²");
                        double height = RequiredNumber(r.Level.Height, "высота верхней надстройки, м");
                        if (area < 8 && height < 2.5) return false;
                    }
                    else if (!IsPublic(r) || r.Role == "roof-exit") return false;
                    else if (RequiredNumber(r.Level.RoofRatio, "суммарная доля технических помещений от кровли") < .15) return false;
                }
                if (r.Level.Kind == "void")
                {
                    if (!IsPublic(r) || IsHigh(r)) return false;
                    return RequiredNumber(r.Level.Height, "высота технического подполья, м") >= 1.8;
                }
                if (IsHigh(r) && Flag(r.Element, "partial-floor") && OneOf(r.Role, "technical-space", "technical-void")) return false;
                return !aboveOnly || Above(r, true);
            }
            private string Part(Record r)
            {
                if (r.Part != "auto") return r.Part;
                if (OneOf(r.Role, "heated", "auxiliary", "residential-common", "entrance") || !string.IsNullOrWhiteSpace(r.Apartment)) return "residential";
                if (OneOf(r.Role, "public", "nonresidential-common", "parking")) return "nonresidential";
                throw new InvalidOperationException("Для роли «" + r.Role + "» задайте обслуживаемую жилую / нежилую часть.");
            }
            private bool Eligible(Record r, Indicator metric)
            {
                if (r.Override.HasValue) return r.Override.Value;
                string include = Mapped(r.Element, "include"); if (Eq(include, "0") || Eq(include, "нет") || Eq(include, "false")) return false;
                if (r.Level == null || !r.Level.Include || r.Level.Kind == "exclude") return false;
                if (r.Role == "unknown") throw new InvalidOperationException("Назначение объекта не определено правилами или параметром функции.");
                string role = r.Role; bool pub = IsPublic(r), gns = (int)metric <= 4;
                if (gns)
                {
                    if (!Above(r)) return false;
                    if (metric == Indicator.GnsResidential && r.BuildingClass != "residential") return false;
                    if (metric == Indicator.GnsNonresidential && r.BuildingClass != "nonresidential") return false;
                    if (metric == Indicator.GnsLivingPart && (r.BuildingClass != "residential" || Part(r) != "residential")) return false;
                    if (metric == Indicator.GnsNonlivingPart && (r.BuildingClass != "residential" || Part(r) != "nonresidential")) return false;
                    if (OneOf(role, "technical-void", "attic", "roof-exit", "roof-vent", "open-stair", "ramp", "canopy", "porch", "pit", "passage", "decoration", "french-balcony", "apartment-family", "ventilated-void", "soil-filled")) return false;
                    if (Config.Method == "new" && OneOf(role, "roof-used", "terrace")) return false;
                    return !OneOf(role, "structure", "envelope", "footprint", "underground-footprint");
                }
                if ((metric == Indicator.GrossAbove || OneOf(metric.ToString(), "Np", "NpResidential", "NpNonresidential", "NnpEmbedded", "NnpSeparate")) && !Above(r)) return false;
                if (metric == Indicator.GrossBelow && Above(r)) return false;
                if (metric == Indicator.NpResidential && r.BuildingClass != "residential") return false;
                if (metric == Indicator.NpNonresidential && r.BuildingClass != "nonresidential") return false;
                if (metric == Indicator.NnpEmbedded || metric == Indicator.NnpSeparate)
                {
                    if (Part(r) != "nonresidential") return false;
                    if (string.IsNullOrWhiteSpace(Config.Parameter("embedded")) || string.IsNullOrWhiteSpace(Config.Parameter("standalone")))
                        throw new InvalidOperationException("Для ННП задайте признаки встроенно-пристроенной части и отдельно стоящего объекта.");
                    if (string.IsNullOrWhiteSpace(Mapped(r.Element, "embedded")) || string.IsNullOrWhiteSpace(Mapped(r.Element, "standalone"))) throw new InvalidOperationException("У нежилого помещения не заполнены признаки встроенно-пристроенной / отдельно стоящей части.");
                    if (!Flag(r.Element, "embedded")) return false;
                    if (Flag(r.Element, "standalone") != (metric == Indicator.NnpSeparate)) return false;
                }
                if (GrossMetric(metric))
                {
                    if (OneOf(role, "structure", "envelope", "footprint", "underground-footprint", "decoration", "stove", "canopy", "porch", "pit", "passage", "french-balcony", "open-stair", "apartment-family")) return false;
                    if (!pub && role == "ramp") return false;
                    if (!pub && OneOf(role, "technical-void", "technical-space", "attic", "roof-exit", "roof-vent", "external-tambour", "tambour", "engineering-shaft")) return false;
                    if (pub && OneOf(role, "attic", "roof-exit")) return false;
                    if (pub && role == "roof-vent") return Dimension(r, "roof-ratio") >= .15;
                    if (OneOf(role, "ventilated-void", "soil-filled")) return false;
                    if (pub && OneOf(role, "technical-void", "technical-space") && Dimension(r, "height", true) < 1.8)
                    { if (string.IsNullOrWhiteSpace(Mapped(r.Element, "service-access"))) throw new InvalidOperationException("Для низкого технического пространства укажите, нужен ли проход обслуживания коммуникаций."); return Flag(r.Element, "service-access"); }
                    if (IsHigh(r) && OneOf(role, "technical-space", "technical-void") && Flag(r.Element, "partial-floor")) return false;
                    if (pub && role == "mezzanine") return r.SingleStorey || Dimension(r, "mezzanine-ratio") > .4;
                    return true;
                }
                if (metric == Indicator.ApartmentsTotal || metric == Indicator.ApartmentsHeated)
                {
                    if (string.IsNullOrWhiteSpace(r.Apartment)) return false;
                    if (metric == Indicator.ApartmentsHeated) return OneOf(role, "heated", "auxiliary", "under-stair", "niche", "arch", "mezzanine");
                    return OneOf(role, "heated", "auxiliary", "under-stair", "niche", "arch", "mezzanine", "balcony", "loggia", "terrace", "veranda", "cold-storage", "tambour");
                }
                if (metric == Indicator.PublicRooms) return role == "public" || r.Part == "nonresidential" && OneOf(role, "external-tambour", "balcony", "loggia", "terrace", "veranda", "stair", "multilight", "opening", "shaft");
                if (metric == Indicator.PublicCalculated)
                    return (pub || role == "public" || r.Part == "nonresidential") && OneOf(role, "public", "heated", "auxiliary", "niche", "arch", "mezzanine");
                return false;
            }
            private double Dimension(Record r, string key, bool length = false)
            { var n = Number(r.Element, key, length); if (!n.HasValue) throw new InvalidOperationException("Нужен параметр: " + Config.Parameters.First(x => x.Key == key).Title); return n.Value; }
            private double Factor(Record r, Indicator metric)
            {
                if (metric != Indicator.ApartmentsTotal) return 1;
                double factor = OneOf(r.Role, "balcony", "terrace") ? .3 : r.Role == "loggia" ? .5 : 1;
                var custom = Number(r.Element, "coefficient"); if (custom.HasValue) factor = custom.Value;
                else if (r.Factor.HasValue) factor = r.Factor.Value;
                if (factor < 0 || factor > 1) throw new InvalidOperationException("Коэффициент должен быть в диапазоне 0..1."); return factor;
            }
            private static IEnumerable<Solid> GeometrySolids(GeometryElement geometry)
            {
                if (geometry == null) yield break;
                foreach (GeometryObject o in geometry)
                {
                    var solid = o as Solid; if (solid != null && solid.Volume > 1e-9) yield return solid;
                    var instance = o as GeometryInstance; if (instance != null) foreach (var s in GeometrySolids(instance.GetInstanceGeometry())) yield return s;
                }
            }
            private static List<Solid> Solids(Element e)
            { return GeometrySolids(e.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false })).ToList(); }
            private const double ArcChordTolerance = .0001 / .3048; // 0.1 mm, reported in every run.
            [ThreadStatic] private static Dictionary<Solid, Tuple<LayeredBody, string>> planarBodies;
            [ThreadStatic] private static Action planarCheckpoint;
            private int planarVolumeCuts;
            private static List<double[]> PolygonRing(CurveLoop loop, ref bool curved)
            {
                var points = new List<double[]>();
                foreach (var curve in loop)
                {
                    if (curve is Line) { var p = curve.GetEndPoint(0); points.Add(new[] { p.X, p.Y }); continue; }
                    var arc = curve as Arc;
                    if (arc == null || Math.Abs(Math.Abs(arc.Normal.Z) - 1) > 1e-10) throw new NotSupportedException("Плоский расчёт поддерживает отрезки и горизонтальные дуги. Для сплайна нужен отдельный расчётный контур.");
                    curved = true;
                    double angle = arc.Length / arc.Radius;
                    double step = 2 * Math.Acos(Math.Max(-1, Math.Min(1, 1 - ArcChordTolerance / arc.Radius)));
                    int count = checked((int)Math.Ceiling(angle / Math.Min(Math.PI / 4, Math.Max(1e-8, step))));
                    if (count > 200000 || points.Count + count > 200000) throw new InvalidOperationException("Слишком сложная дуга для допуска 0,1 мм. Разделите расчётный контур.");
                    for (int i = 0; i < count; i++)
                    {
                        XYZ p;
                        if (arc.IsBound) p = arc.Evaluate((double)i / count, true);
                        else { double a = 2 * Math.PI * i / count; p = arc.Center + arc.Radius * (Math.Cos(a) * arc.XDirection + Math.Sin(a) * arc.YDirection); }
                        points.Add(new[] { p.X, p.Y });
                    }
                }
                return points;
            }
            private static PlanarRegion FaceRegion(PlanarFace face, ref bool curved)
            {
                var rings = new List<List<double[]>>();
                foreach (var loop in face.GetEdgesAsCurveLoops()) rings.Add(PolygonRing(loop, ref curved));
                return PlanarRegion.FromRings(rings);
            }
            private sealed class HorizontalCaps
            { internal PlanarRegion Up = PlanarRegion.Empty, Down = PlanarRegion.Empty; }
            private static bool TryLayeredBody(Solid solid, out LayeredBody body, out string reason)
            {
                body = null; reason = null; if (solid == null) { body = new LayeredBody(); return true; }
                Tuple<LayeredBody, string> cached;
                if (planarBodies != null && planarBodies.TryGetValue(solid, out cached)) { body = cached.Item1; reason = cached.Item2; return body != null; }
                try
                {
                    planarCheckpoint?.Invoke();
                    if (solid.Volume <= 0 || double.IsNaN(solid.Volume) || double.IsInfinity(solid.Volume)) throw new NotSupportedException("Исходное тело не имеет положительного конечного объёма.");
                    var caps = new SortedDictionary<double, HorizontalCaps>(); bool curved = false;
                    foreach (Face face in solid.Faces)
                    {
                        var plane = face as PlanarFace;
                        if (plane != null)
                        {
                            if (Math.Abs(plane.FaceNormal.Z) < 1e-10) continue;
                            if (Math.Abs(Math.Abs(plane.FaceNormal.Z) - 1) > 1e-10) throw new NotSupportedException("Есть наклонная грань; точное разложение на вертикальные участки неприменимо.");
                            double z = Math.Round(plane.Origin.Z, 8); HorizontalCaps cap;
                            if (!caps.TryGetValue(z, out cap)) caps[z] = cap = new HorizontalCaps();
                            var region = FaceRegion(plane, ref curved);
                            if (plane.FaceNormal.Z > 0) cap.Up = PlanarRegion.Combine(cap.Up, region, CT.ctUnion);
                            else cap.Down = PlanarRegion.Combine(cap.Down, region, CT.ctUnion);
                        }
                        else
                        {
                            var cylinder = face as CylindricalFace;
                            if (cylinder == null || Math.Abs(Math.Abs(cylinder.Axis.Normalize().Z) - 1) > 1e-10)
                                throw new NotSupportedException("Есть непризматическая криволинейная грань; используется объёмная геометрия Revit.");
                        }
                    }
                    if (caps.Count < 2) throw new NotSupportedException("Не найдены горизонтальные границы высотных участков.");
                    var candidate = new LayeredBody { Curved = curved }; var active = PlanarRegion.Empty; var levels = caps.Keys.ToList();
                    for (int i = 0; i < levels.Count; i++)
                    {
                        planarCheckpoint?.Invoke(); var cap = caps[levels[i]];
                        active = PlanarRegion.Combine(active, cap.Down, CT.ctUnion);
                        active = PlanarRegion.Combine(active, cap.Up, CT.ctDifference);
                        if (i + 1 < levels.Count && !active.IsEmpty) candidate.Layers.Add(new PlanarLayer { Bottom = levels[i], Top = levels[i + 1], Region = active });
                    }
                    if (!active.IsEmpty || candidate.Layers.Count == 0) throw new NotSupportedException("Горизонтальные грани не образуют замкнутый набор вертикальных участков.");
                    double tolerance = Math.Max(1e-7, solid.Volume * 1e-8) + candidate.Layers.Sum(l => l.Region.Perimeter * (curved ? ArcChordTolerance : 2 / PlanarRegion.Scale) * (l.Top - l.Bottom) * 4 + l.Region.Area * 2e-8);
                    if (Math.Abs(candidate.Volume - solid.Volume) > tolerance) throw new NotSupportedException("Разложение не сохранило объём в пределах допуска; исходное тело не заменено приближённой призмой.");
                    body = candidate;
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { reason = ex.Message; }
                if (planarBodies != null) planarBodies[solid] = Tuple.Create(body, reason);
                return body != null;
            }
            private static LayeredBody RequireLayeredBody(Solid solid)
            { LayeredBody body; string reason; if (!TryLayeredBody(solid, out body, out reason)) throw new InvalidOperationException("Плоское вычитание: " + reason); return body; }
            private static Solid BuildPlanarSolid(PlanarRegion region, double bottom, double top, bool curved = false)
            {
                if (region == null || region.IsEmpty || top <= bottom) return null;
                double expected = region.Area * (top - bottom); Solid shape = null; var failures = new List<string>();
                // Revit's curve-construction tolerance is larger than its geometric precision.
                // Build small edges in local enlarged coordinates, then apply an exact uniform
                // inverse transform. No vertices are removed or shifted relative to each other.
                foreach (double scale in new[] { 1.0, 1000.0 })
                {
                    try
                    {
                        planarCheckpoint?.Invoke(); var origin = region.Paths[0][0];
                        double x = origin.X / PlanarRegion.Scale, y = origin.Y / PlanarRegion.Scale;
                        var loops = new List<CurveLoop>();
                        foreach (var path in region.Paths)
                        {
                            var curves = new List<Curve>();
                            for (int i = 0; i < path.Count; i++)
                            {
                                var p = path[i]; var q = path[(i + 1) % path.Count];
                                curves.Add(Line.CreateBound(new XYZ((p.X - origin.X) / PlanarRegion.Scale * scale, (p.Y - origin.Y) / PlanarRegion.Scale * scale, 0),
                                    new XYZ((q.X - origin.X) / PlanarRegion.Scale * scale, (q.Y - origin.Y) / PlanarRegion.Scale * scale, 0)));
                            }
                            loops.Add(CurveLoop.Create(curves));
                        }
                        var local = Extrude(loops, (top - bottom) * scale);
                        var transform = Transform.Identity; transform.BasisX = XYZ.BasisX / scale; transform.BasisY = XYZ.BasisY / scale; transform.BasisZ = XYZ.BasisZ / scale; transform.Origin = new XYZ(x, y, bottom);
                        var candidate = SolidUtils.CreateTransformed(local, transform);
                        if (candidate == null || !FinitePositiveVolume(candidate.Volume) || Math.Abs(candidate.Volume - expected) > Math.Max(1e-7, expected * 1e-7))
                            throw new InvalidOperationException("Revit изменил площадь / объём при построении плоского результата.");
                        shape = candidate; break;
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { failures.Add("Масштаб " + scale.ToString(CultureInfo.InvariantCulture) + ": " + ex.Message); }
                }
                if (shape == null) throw new InvalidOperationException("Плоский результат вычислен, но Revit не построил тело. Малые участки не удалены. " + string.Join(" | ", failures));
                if (planarBodies != null)
                { var body = new LayeredBody { Curved = curved }; body.Layers.Add(new PlanarLayer { Bottom = bottom, Top = top, Region = region }); planarBodies[shape] = Tuple.Create(body, (string)null); }
                return shape;
            }
            private static bool FinitePositiveVolume(double value) { return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value); }
            private static List<Solid> BuildLayeredSolids(LayeredBody body)
            {
                var result = new List<Solid>();
                foreach (var layer in body.Layers)
                    foreach (var paths in layer.Region.Components())
                    {
                        planarCheckpoint?.Invoke(); var region = new PlanarRegion(paths);
                        try { result.Add(BuildPlanarSolid(region, layer.Bottom, layer.Top, body.Curved)); }
                        catch (System.OperationCanceledException) { throw; }
                        catch (Exception original)
                        {
                            // Touching hole/outer boundaries can be valid polygons yet invalid
                            // Revit extrusion profiles. Split at every X vertex into simple strips.
                            var xs = paths.SelectMany(p => p).Select(p => p.X).Distinct().OrderBy(v => v).ToList();
                            long minY = paths.SelectMany(p => p).Min(p => p.Y), maxY = paths.SelectMany(p => p).Max(p => p.Y);
                            var strips = new List<Solid>(); double total = 0;
                            try
                            {
                                for (int i = 0; i + 1 < xs.Count; i++)
                                {
                                    planarCheckpoint?.Invoke(); var box = new PlanarRegion(new List<List<CP>> { new List<CP> { new CP(xs[i], minY), new CP(xs[i + 1], minY), new CP(xs[i + 1], maxY), new CP(xs[i], maxY) } });
                                    var clipped = PlanarRegion.Combine(region, box, CT.ctIntersection); total += clipped.Area;
                                    foreach (var part in clipped.Components()) strips.Add(BuildPlanarSolid(new PlanarRegion(part), layer.Bottom, layer.Top, body.Curved));
                                }
                                if (Math.Abs(total - region.Area) > Math.Max(1e-8, region.Perimeter * 4 / PlanarRegion.Scale)) throw new InvalidOperationException("Разбиение на полосы не сохранило площадь.");
                                if (strips.Count == 0) throw new InvalidOperationException("Разбиение на полосы вернуло пустой результат.");
                                result.AddRange(strips);
                            }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { throw new InvalidOperationException(original.Message + " | Разбиение плоского результата: " + ex.Message, ex); }
                        }
                    }
                return result;
            }
            private static Solid FlatBoolean(Solid a, Solid b, CT operation)
            {
                var aa = RequireLayeredBody(a); var bb = RequireLayeredBody(b);
                if (aa.Layers.Count != 1 || bb.Layers.Count != 1) throw new InvalidOperationException("Операция площади требует одного плоского слоя у каждого контура.");
                var x = aa.Layers[0]; var y = bb.Layers[0];
                if (Math.Abs(x.Bottom - y.Bottom) > 1e-7 || Math.Abs(x.Top - y.Top) > 1e-7) throw new InvalidOperationException("Контуры площади лежат на разных расчётных плоскостях.");
                planarCheckpoint?.Invoke();
                var region = PlanarRegion.Combine(x.Region, y.Region, operation);
                if (operation == CT.ctDifference)
                {
                    var overlap = PlanarRegion.Combine(x.Region, y.Region, CT.ctIntersection);
                    double tolerance = Math.Max(1e-8, (x.Region.Perimeter + y.Region.Perimeter) * 4 / PlanarRegion.Scale);
                    if (Math.Abs(x.Region.Area - region.Area - overlap.Area) > tolerance) throw new InvalidOperationException("Нарушен баланс площадей плоского вычитания.");
                }
                return BuildPlanarSolid(region, x.Bottom, x.Top, aa.Curved || bb.Curved);
            }
            private bool TryLayeredDifference(Solid a, Solid b, out List<Solid> result, out string reason)
            {
                result = null; reason = null; LayeredBody aa, bb; string aReason, bReason;
                bool aOk = TryLayeredBody(a, out aa, out aReason), bOk = TryLayeredBody(b, out bb, out bReason);
                if (!aOk || !bOk) { reason = "A: " + (aReason ?? "вертикальные участки") + "; B: " + (bReason ?? "вертикальные участки"); return false; }
                try
                {
                    var difference = LayeredBody.Combine(aa, bb, CT.ctDifference);
                    var overlap = LayeredBody.Combine(aa, bb, CT.ctIntersection);
                    double tolerance = Math.Max(1e-7, (aa.Layers.Concat(bb.Layers).Sum(l => l.Region.Perimeter * (l.Top - l.Bottom))) * 4 / PlanarRegion.Scale);
                    if (Math.Abs(aa.Volume - difference.Volume - overlap.Volume) > tolerance) throw new InvalidOperationException("Нарушен баланс объёмов высотных участков.");
                    result = BuildLayeredSolids(difference); planarVolumeCuts++; return true;
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { reason = ex.Message; return false; }
            }
            private static List<Solid> VolumeIntersections(Solid a, Solid b)
            {
                if (a == null || b == null || !Overlaps(Bounds(a), Bounds(b))) return new List<Solid>();
                LayeredBody aa, bb; string reason;
                if (TryLayeredBody(a, out aa, out reason) && TryLayeredBody(b, out bb, out reason))
                    return BuildLayeredSolids(LayeredBody.Combine(aa, bb, CT.ctIntersection));
                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(a, b, BooleanOperationsType.Intersect);
                return intersection == null || intersection.Volume < 1e-9 ? new List<Solid>() : new List<Solid> { intersection };
            }
            private List<Solid> DecomposeVolumeInput(Solid solid, Record record, string metric)
            {
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason) && body.Layers.Count > 1)
                {
                    try { return BuildLayeredSolids(body); }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("PLANAR_REBUILD", "Предупреждение", "Не удалось представить исходное тело отдельными вертикальными участками; используется исходная геометрия. " + ex.Message, record, metric); }
                }
                return new List<Solid> { solid };
            }
            private static string BooleanObject(Record record)
            {
                return record == null ? "не определён" : record.Source.Name + "; ID " + IDHelper.ElIdValue(record.Element.Id) + "; UniqueId " + record.Element.UniqueId +
                    "; корпус " + record.Building + "; этаж " + (record.Level?.Name ?? "не определён");
            }

            // XY index for planar areas; XYZ index for volumes. Bounds are only a broad-phase filter.
            private sealed class PlanIndex
            {
                private readonly bool spatial;
                private readonly Dictionary<Tuple<int, int, int>, List<int>> cells = new Dictionary<Tuple<int, int, int>, List<int>>();
                private readonly List<int> large = new List<int>();
                private readonly List<Tuple<Solid, double[]>> entries = new List<Tuple<Solid, double[]>>();
                internal PlanIndex(bool spatial = false) { this.spatial = spatial; }
                private static int Cell(double value) { return checked((int)Math.Floor(value / 32.0)); }
                private int[] Range(double[] b) { return new[] { Cell(b[0]), Cell(b[3]), Cell(b[1]), Cell(b[4]), spatial ? Cell(b[2]) : 0, spatial ? Cell(b[5]) : 0 }; }
                private static bool Large(int[] r) { return ((double)r[1] - r[0] + 1) * ((double)r[3] - r[2] + 1) * ((double)r[5] - r[4] + 1) > 4096; }
                private readonly Dictionary<Solid, Record> owners = new Dictionary<Solid, Record>();
                internal Record OwnerOf(Solid solid) { Record owner; return owners.TryGetValue(solid, out owner) ? owner : null; }
                internal void Add(Solid solid, Record owner = null)
                {
                    if (solid == null || solid.Volume < 1e-9) return;
                    if (owner != null) owners[solid] = owner;
                    var b = Bounds(solid); var r = Range(b); int id = entries.Count; entries.Add(Tuple.Create(solid, b));
                    if (Large(r)) { large.Add(id); return; }
                    for (int x = r[0]; x <= r[1]; x++) for (int y = r[2]; y <= r[3]; y++) for (int z = r[4]; z <= r[5]; z++)
                            { var key = Tuple.Create(x, y, z); List<int> bucket; if (!cells.TryGetValue(key, out bucket)) cells[key] = bucket = new List<int>(); bucket.Add(id); }
                }
                internal IEnumerable<Solid> Query(Solid solid)
                {
                    if (solid == null || solid.Volume < 1e-9) yield break;
                    var b = Bounds(solid); var r = Range(b); var found = new HashSet<int>(large);
                    if (Large(r)) { for (int i = 0; i < entries.Count; i++) found.Add(i); }
                    else for (int x = r[0]; x <= r[1]; x++) for (int y = r[2]; y <= r[3]; y++) for (int z = r[4]; z <= r[5]; z++)
                                { List<int> bucket; if (cells.TryGetValue(Tuple.Create(x, y, z), out bucket)) foreach (int id in bucket) found.Add(id); }
                    foreach (int id in found.OrderBy(i => i)) if (Overlaps(b, entries[id].Item2)) yield return entries[id].Item1;
                }
            }
            private static double[] Bounds(Solid solid)
            {
                var box = solid.GetBoundingBox(); var result = new[] { double.PositiveInfinity, double.PositiveInfinity, double.PositiveInfinity, double.NegativeInfinity, double.NegativeInfinity, double.NegativeInfinity };
                for (int i = 0; i < 8; i++)
                {
                    var p = box.Transform.OfPoint(new XYZ((i & 1) == 0 ? box.Min.X : box.Max.X, (i & 2) == 0 ? box.Min.Y : box.Max.Y, (i & 4) == 0 ? box.Min.Z : box.Max.Z));
                    result[0] = Math.Min(result[0], p.X); result[1] = Math.Min(result[1], p.Y); result[2] = Math.Min(result[2], p.Z);
                    result[3] = Math.Max(result[3], p.X); result[4] = Math.Max(result[4], p.Y); result[5] = Math.Max(result[5], p.Z);
                }
                return result;
            }
            private static bool Overlaps(double[] a, double[] b)
            { return a[0] < b[3] && a[3] > b[0] && a[1] < b[4] && a[4] > b[1] && a[2] < b[5] && a[5] > b[2]; }
            private static Solid Union(Solid a, Solid b)
            { if (a == null || a.Volume < 1e-9) return b; if (b == null || b.Volume < 1e-9) return a; return FlatBoolean(a, b, CT.ctUnion); }
            private static Solid Subtract(Solid a, Solid b)
            { if (a == null || b == null || !Overlaps(Bounds(a), Bounds(b))) return a; return FlatBoolean(a, b, CT.ctDifference); }
            private static Solid Intersect(Solid a, Solid b)
            { if (a == null || b == null || !Overlaps(Bounds(a), Bounds(b))) return null; return FlatBoolean(a, b, CT.ctIntersection); }
            private static Curve Flat(Curve c, double z = 0)
            {
                if (c is Line) { var p = c.GetEndPoint(0); var q = c.GetEndPoint(1); return Line.CreateBound(new XYZ(p.X, p.Y, z), new XYZ(q.X, q.Y, z)); }
                if (c is Arc)
                {
                    var a = (Arc)c; if (Math.Abs(a.Normal.Z) < .999999) throw new InvalidOperationException("Наклонная дуга требует явного горизонтального расчётного контура.");
                    return c.CreateTransformed(Transform.CreateTranslation(new XYZ(0, 0, z - c.GetEndPoint(0).Z)));
                }
                if (Math.Abs(c.GetEndPoint(0).Z - c.GetEndPoint(1).Z) > 1e-7) throw new InvalidOperationException("Неплоская кривая расчётного контура.");
                return c.CreateTransformed(Transform.CreateTranslation(new XYZ(0, 0, z - c.GetEndPoint(0).Z)));
            }
            private static CurveLoop FlatLoop(IEnumerable<Curve> curves, double z = 0) { return CurveLoop.Create(curves.Select(c => Flat(c, z)).ToList()); }
            private static Solid Extrude(IList<CurveLoop> loops, double height = 1)
            { return GeometryCreationUtilities.CreateExtrusionGeometry(loops, XYZ.BasisZ, height); }
            private static IList<CurveLoop> BottomLoops(Solid solid, double z = 0)
            {
                var faces = solid.Faces.Cast<Face>().OfType<PlanarFace>().Where(f => f.FaceNormal.Z < -.999999).ToList();
                if (faces.Count != 1) throw new InvalidOperationException("Для обводки требуется один связный плоский контур; геометрия разбивается по граням.");
                return faces[0].GetEdgesAsCurveLoops().Select(l => FlatLoop(l, z)).ToList();
            }
            private static Solid Projection(Solid solid)
            {
                LayeredBody body; string reason; if (TryLayeredBody(solid, out body, out reason)) return BuildPlanarSolid(body.Projection(), 0, 1, body.Curved);
                Solid result = null;
                foreach (Face f in solid.Faces)
                {
                    var p = f as PlanarFace; if (p == null)
                    { var cylinder = f as CylindricalFace; if (cylinder != null && Math.Abs(cylinder.Axis.Z) > .999999) continue; throw new InvalidOperationException("Проекция криволинейной поверхности требует явной зоны / семейства расчётного контура."); }
                    if (p.FaceNormal.Z <= 1e-7) continue;
                    result = Union(result, Extrude(p.GetEdgesAsCurveLoops().Select(l => FlatLoop(l)).ToList()));
                }
                return result ?? throw new InvalidOperationException("Не найден замкнутый контур горизонтальной проекции.");
            }
            private static Solid Section(Solid solid, double z)
            {
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason))
                { var profile = PlanarRegion.Empty; foreach (var l in body.Layers.Where(l => l.Bottom <= z + 1e-8 && l.Top >= z - 1e-8)) profile = PlanarRegion.Combine(profile, l.Region, CT.ctUnion); return BuildPlanarSolid(profile, 0, 1, body.Curved); }
                Solid atBoundary = null;
                foreach (PlanarFace f in solid.Faces.Cast<Face>().OfType<PlanarFace>())
                    if (Math.Abs(f.FaceNormal.Z) > .999999 && Math.Abs(f.Origin.Z - z) < 1e-7)
                        atBoundary = Union(atBoundary, Extrude(f.GetEdgesAsCurveLoops().Select(l => FlatLoop(l)).ToList()));
                Solid cut = Half(solid, z, false);
                if (cut == null || cut.Volume < 1e-9) return atBoundary;
                Solid result = atBoundary;
                foreach (PlanarFace f in cut.Faces.Cast<Face>().OfType<PlanarFace>())
                    if (f.FaceNormal.Z > .999999 && Math.Abs(f.Origin.Z - z) < 1e-6)
                        result = Union(result, Extrude(f.GetEdgesAsCurveLoops().Select(l => FlatLoop(l)).ToList()));
                return result;
            }
            private static Solid Half(Solid solid, double z, bool above)
            {
                if (solid == null || solid.Volume < 1e-9) return null;
                var box = solid.GetBoundingBox(); var elevations = Enumerable.Range(0, 8).Select(i => box.Transform.OfPoint(new XYZ((i & 1) == 0 ? box.Min.X : box.Max.X, (i & 2) == 0 ? box.Min.Y : box.Max.Y, (i & 4) == 0 ? box.Min.Z : box.Max.Z)).Z).ToList();
                if (above) { if (elevations.Max() <= z) return null; if (elevations.Min() >= z) return solid; }
                else { if (elevations.Min() >= z) return null; if (elevations.Max() <= z) return solid; }
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason) && body.Layers.Count == 1)
                { var l = body.Layers[0]; return BuildPlanarSolid(l.Region, above ? Math.Max(z, l.Bottom) : l.Bottom, above ? l.Top : Math.Min(z, l.Top), body.Curved); }
                return BooleanOperationsUtils.CutWithHalfSpace(solid, Plane.CreateByNormalAndOrigin(above ? XYZ.BasisZ : -XYZ.BasisZ, new XYZ(0, 0, z)));
            }
            private static bool Inside(Solid solid, XYZ point)
            {
                using (var hit = solid.IntersectWithCurve(Line.CreateBound(new XYZ(point.X, point.Y, .25), new XYZ(point.X, point.Y, .75)), new SolidCurveIntersectionOptions())) return hit.SegmentCount > 0;
            }
            private Solid SpatialPlan(Record r, string boundary, string metric = "")
            {
                var spatial = (SpatialElement)r.Element;
                var options = new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = boundary == "net" ? SpatialElementBoundaryLocation.Finish : SpatialElementBoundaryLocation.Center };
                var boundaries = spatial.GetBoundarySegments(options);
                if (boundaries == null || boundaries.Count == 0) throw new InvalidOperationException("Нет замкнутых границ помещения / зоны.");
                var loops = boundaries.Select(b => FlatLoop(b.Select(x => x.GetCurve()))).ToList();
                if (!(spatial is Area) && boundary != "net")
                {
                    var centre = Extrude(loops); var modified = new List<CurveLoop>();
                    for (int li = 0; li < loops.Count; li++)
                    {
                        var distances = new List<double>(); bool changed = false;
                        foreach (var segment in boundaries[li])
                        {
                            var boundaryElement = r.Element.Document.GetElement(segment.ElementId); var wall = boundaryElement as Wall; var wallTransform = Transform.Identity; double d = 0;
                            var boundaryLink = boundaryElement as RevitLinkInstance;
                            if (boundaryLink != null) { wall = boundaryLink.GetLinkDocument()?.GetElement(segment.LinkElementId) as Wall; wallTransform = boundaryLink.GetTotalTransform(); if (wall == null) throw new InvalidOperationException("Граница связана со связью, но исходная стена недоступна. Задайте расчётный контур зоной."); }
                            if (wall != null && wall.WallType.Function == WallFunction.Exterior)
                            {
                                if (wall.WallType.Kind != WallKind.Basic) throw new InvalidOperationException("Нестандартная наружная стена: задайте внешний / внутренний контур зоной.");
                                BuiltInParameter sectionParameter; if (Enum.TryParse("WALL_CROSS_SECTION", out sectionParameter))
                                { var crossSection = wall.get_Parameter(sectionParameter); if (crossSection != null && crossSection.AsInteger() != 0) throw new InvalidOperationException("Наклонная / коническая наружная стена требует явного контура на нормативной высоте обмера."); }
                                var c = Flat(segment.GetCurve()); var p = c.Evaluate(.5, true); var tangent = c.ComputeDerivatives(.5, true).BasisX.Normalize();
                                var right = tangent.CrossProduct(XYZ.BasisZ).Normalize(); double sign = Inside(centre, p + right * .002) ? -1 : 1;
                                double width = wall.Width / 2;
                                if (boundary == "outer" && Config.Method == "new")
                                {
                                    var compound = wall.WallType.GetCompoundStructure(); if (compound != null)
                                    {
                                        var layers = compound.GetLayers(); int first = compound.GetFirstCoreLayerIndex(); if (first < 0) throw new InvalidOperationException("У наружной стены не задан несущий слой для нового ГНС.");
                                        double finish = 0; bool exterior = wallTransform.OfVector(wall.Orientation).DotProduct(right * sign) > 0;
                                        if (exterior) { for (int i = 0; i < first; i++) finish += layers[i].Width; }
                                        else { for (int i = compound.GetLastCoreLayerIndex() + 1; i < layers.Count; i++) finish += layers[i].Width; }
                                        width -= finish;
                                    }
                                }
                                d = sign * (boundary == "outer" ? width : -wall.Width / 2); changed |= Math.Abs(d) > 1e-9;
                            }
                            else if (segment.ElementId == ElementId.InvalidElementId && boundary == "outer")
                                Notice("BOUNDARY_SEPARATOR", "Предупреждение", "Часть границы не связана с наружной стеной. Проверьте внешний контур на расчётном виде.", r, metric);
                            distances.Add(d);
                        }
                        modified.Add(changed ? CurveLoop.CreateViaOffset(loops[li], distances, XYZ.BasisZ) : loops[li]);
                    }
                    loops = modified;
                }
                var local = Extrude(loops);
                var horizontal = Transform.CreateTranslation(new XYZ(0, 0, -r.Source.Transform.Origin.Z)).Multiply(r.Source.Transform);
                return SolidUtils.CreateTransformed(local, horizontal);
            }
            private Solid Plan(Record r, Indicator metric)
            {
                Progress("Геометрия: " + MetricLabels[metric.ToString()] + " / " + r.Source.Name + " / " + r.Level?.Name + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                string boundary = (int)metric <= 4 ? "outer" : GrossMetric(metric) ? "gross" : "net";
                if (boundary == "outer" && r.Element is SpatialElement && !(r.Element is Area)) Notice("EXTERIOR_FUNCTION", "Предупреждение", "Автоматический внешний контур использует функцию наружных стен и их ядро. Проверьте маркировку наружных стен и полноту помещений на служебном плане.", r, metric.ToString());
                if (r.Element is Area) Notice("AREA_CONTOUR", "Предупреждение", "Используется готовая граница зоны. Проверьте соответствие схемы: наружный контур для ГНС, внутренний для общей / наземной площади, чистовой для квартир и помещений.", r, metric.ToString());
                if (boundary == "net" && !(r.Element is SpatialElement) && !(r.Element is Opening)) Notice("FAMILY_NET_CONTOUR", "Предупреждение", "Для семейства используется горизонтальная проекция. Проверьте чистовой контур и исключение толщины ограждений; при необходимости используйте отдельную зону.", r, metric.ToString());
                string key = r.Key + "/" + boundary; Solid shape;
                if (!shapes.TryGetValue(key, out shape))
                {
                    if (r.Element is SpatialElement) shape = SpatialPlan(r, boundary, metric.ToString());
                    else if (r.Element is Opening)
                    {
                        var opening = (Opening)r.Element; var curves = opening.BoundaryCurves.Cast<Curve>().ToList();
                        if (curves.Count == 0) throw new InvalidOperationException("У проёма нет горизонтальной границы. Назначьте расчётную зону.");
                        var transformed = curves.Select(c => c.CreateTransformed(r.Source.Transform)); shape = Extrude(new List<CurveLoop> { FlatLoop(transformed) });
                    }
                    else
                    {
                        shape = null; foreach (var solid in Solids(r.Element)) shape = Union(shape, Projection(SolidUtils.CreateTransformed(solid, r.Source.Transform)));
                        if (shape == null) throw new InvalidOperationException("Нет замкнутой геометрии объекта.");
                    }
                    shapes[key] = shape;
                }
                if (boundary != "outer")
                {
                    double min = 0;
                    bool residentialRoom = metric == Indicator.ApartmentsTotal || metric == Indicator.ApartmentsHeated || !IsPublic(r) && r.Part != "nonresidential" && r.Role != "public";
                    if (residentialRoom && r.Role == "under-stair") min = 1.600001;
                    if (residentialRoom && r.Role == "niche" && Dimension(r, "height", true) < 2) return null;
                    if (residentialRoom && r.Role == "arch" && Dimension(r, "width", true) < 2) return null;
                    var slope = Number(r.Element, "slope");
                    if (slope.HasValue)
                    {
                        if (!IsPublic(r) && r.Role != "public" && r.Part != "nonresidential")
                        {
                            if (r.Element is Area && r.Override == true) { Notice("MANSARD_MANUAL", "Предупреждение", "Использован явно включённый контур жилой мансарды. Проверьте согласованное обоснование высотной границы.", r, metric.ToString()); return shape; }
                            throw new InvalidOperationException("Для жилой мансарды в исходном ТЗ неоднозначно заданы пороги высоты. Подготовьте зону учитываемой части и отдельное правило включения с обоснованием.");
                        }
                        min = metric == Indicator.PublicCalculated ? 1.5 : PublicMansardHeight(slope.Value);
                    }
                    if (min > 0 && r.Element is SpatialElement && !(r.Element is Area))
                    {
                        var volume = SpatialVolume(r, boundary == "gross"); var section = Section(volume, r.Z + min / .3048);
                        shape = Intersect(shape, section);
                    }
                    else if (min > 0 && r.Element is Area)
                        Notice("AREA_HEIGHT", "Предупреждение", "Зона должна быть заранее обрезана по нормативной высоте: " + min.ToString("0.###") + " м.", r, metric.ToString());
                }
                return shape;
            }
            public static double PublicMansardHeight(double angle)
            { if (angle < 0 || angle > 90) throw new ArgumentOutOfRangeException("angle"); return angle <= 30 ? 1.5 : angle <= 45 ? 1.5 - (angle - 30) * .4 / 15 : angle <= 60 ? 1.1 - (angle - 45) * .6 / 15 : .5; }
            private Solid SpatialVolume(Record r, bool centre = false)
            {
                string key = r.Key + (centre ? "/volume-centre" : "/volume"); Solid cached; if (shapes.TryGetValue(key, out cached)) return cached;
                if (r.Element is Area) throw new InvalidOperationException("Зона не содержит высоту и не определяет строительный объём. Нужны Rooms/Spaces и оболочка либо замкнутый расчётный объём.");
                var localKey = Tuple.Create(r.Element, centre); Solid local;
                if (!localSpatialVolumes.TryGetValue(localKey, out local))
                {
                    Progress("Геометрия помещения / пространства: " + r.Source.Name + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    local = Measure("Геометрия Rooms/Spaces", r, "", () =>
                    {
                        var calculatorKey = Tuple.Create(r.Element.Document, centre); SpatialElementGeometryCalculator calculator;
                        if (!spatialCalculators.TryGetValue(calculatorKey, out calculator))
                            spatialCalculators[calculatorKey] = calculator = new SpatialElementGeometryCalculator(r.Element.Document, new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = centre ? SpatialElementBoundaryLocation.Center : SpatialElementBoundaryLocation.Finish });
                        return calculator.CalculateSpatialElementGeometry((SpatialElement)r.Element).GetGeometry();
                    });
                    localSpatialVolumes.Add(localKey, local);
                }
                else spatialCacheHits++;
                cached = SolidUtils.CreateTransformed(local, r.Source.Transform); shapes[key] = cached; return cached;
            }
            public Run Calculate(Action<string> progress = null, bool? createViews = null)
            {
                reportProgress = progress; var previous = Last;
                using (var group = new TransactionGroup(doc, "ТЭП: расчёт и проверочные виды"))
                {
                    group.Start();
                    try { var run = CalculateCore(progress, createViews); Progress("Завершение расчёта..."); if (group.Assimilate() != TransactionStatus.Committed) throw new InvalidOperationException("Revit отменил группу транзакций расчёта."); Last = run; return run; }
                    catch { if (group.GetStatus() == TransactionStatus.Started) group.RollBack(); Last = previous; throw; }
                    finally { createViewsForRun = null; planarBodies = null; planarCheckpoint = null; reportProgress = null; parameterCache.Clear(); phaseCache.Clear(); ClearVolumeCaches(); }
                }
            }
            private Run CalculateCore(Action<string> progress, bool? createViews)
            {
                planarBodies = new Dictionary<Solid, Tuple<LayeredBody, string>>(); progressMessage = "Подготовка расчёта контуров..."; planarCheckpoint = () => reportProgress?.Invoke(progressMessage); planarVolumeCuts = 0;
                parameterCache.Clear(); phaseCache.Clear(); ClearVolumeCaches(); timings.Clear(); slowOperations.Clear(); volumeCacheHits = 0; spatialCacheHits = 0; booleanTouchSkips = 0; booleanSplitRecoveries = 0; booleanIntersectionRecoveries = 0; booleanNormalizedRecoveries = 0;
                SnapshotSources(); Normalize(Config); notices.Clear(); shapes.Clear(); appliedCorrections.Clear();
                if (createViews.HasValue) Config.CreateViews = createViews.Value;
                createViewsForRun = Config.CreateViews;
                current = new Run { Author = app.Application.Username, Method = Choices("method").First(x => x.Key == Config.Method).Label };
                current.CreateViews = ViewsEnabled;
                current.Issues.AddRange(startupIssues.Where(i => string.IsNullOrEmpty(i.Source) || Sources.Any(s => s.Name == i.Source && s.Mode != "exclude")));
                foreach (var c in Config.Corrections)
                { if (string.IsNullOrWhiteSpace(c.Author)) c.Author = current.Author; if (string.IsNullOrWhiteSpace(c.Date)) c.Date = current.Date; }
                current.Configuration = Serialize(Config);
                string zeroError = null, groundError = null;
                try { zero = Datum(Config.ZeroMode, Config.ZeroValue, Config.ZeroLevel, Config.ZeroParameter, false); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { zeroError = ex.Message; }
                try { ground = Datum(Config.GroundMode, Config.GroundValue, Config.GroundLevel, Config.GroundParameter, true); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { groundError = ex.Message; }
                progress?.Invoke("Чтение источников и классификация объектов..."); var records = Collect();
                foreach (var metric in Config.Metrics.Where(x => x.Enabled))
                {
                    progress?.Invoke("Расчёт: " + metric.Name); var indicator = (Indicator)Enum.Parse(typeof(Indicator), metric.Key);
                    try
                    {
                        if (zeroError != null && metric.Key.StartsWith("Volume")) throw new InvalidOperationException(zeroError);
                        if (groundError != null && ((int)indicator <= 5 || OneOf(metric.Key, "GrossAbove", "GrossBelow", "NnpEmbedded", "NnpSeparate", "Np", "NpResidential", "NpNonresidential", "ParkingCount", "Storeys"))) throw new InvalidOperationException(groundError);
                        var adjusted = new List<Record>();
                        foreach (var original in records)
                        {
                            Progress("Правила: " + metric.Name + " - " + (adjusted.Count + 1) + " / " + records.Count);
                            var r = original.Copy(); try { ApplyRules(r, metric.Key); adjusted.Add(r); }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { Notice("RULE_CONFLICT", "Ошибка", ex.Message, r, metric.Key); }
                        }
                        adjusted = Representations(adjusted, metric);
                        foreach (var group in adjusted.GroupBy(r => r.Building))
                        { bool single = group.Where(r => r.Level != null && r.Level.Include && r.Role != "mezzanine").Select(r => Math.Round(r.Z, 6)).Distinct().Count() == 1; foreach (var r in group) r.SingleStorey = single; }
                        if (metric.ContourMode != "current") CalculateContours(adjusted, indicator, metric);
                        else if (indicator == Indicator.Storeys || indicator == Indicator.Floors) CalculateFloors(adjusted, indicator);
                        else if (indicator == Indicator.ApartmentsCount || indicator == Indicator.ParkingCount) CalculateCounts(adjusted, indicator);
                        else if (indicator == Indicator.Volume || indicator == Indicator.VolumeAbove || indicator == Indicator.VolumeBelow) CalculateVolumes(adjusted, indicator);
                        else if (indicator == Indicator.Footprint) CalculateFootprints(adjusted);
                        else CalculateAreas(adjusted, indicator);
                        ApplyDeltas(indicator);
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { current.Issue("METRIC_FAILED", "Ошибка", ex.Message, metric.Key); }
                    current.Issue("CALCULATION_PATH", "Информация", "Геометрия " + GeometryVersion + "; источник контура: " + Choices("contour").First(c => c.Key == metric.ContourMode).Label +
                        ". Плоские операции: сетка 0,001 мм; отклонение хорд дуг не более 0,1 мм. Непризматические тела сохраняют обработку Revit.", metric.Key);
                    Summarize(metric);
                }
                foreach (var c in Config.Corrections.Where(c => !appliedCorrections.Contains(c)))
                {
                    if (c.Metric != "all" && !Config.Metrics.Any(m => m.Key == c.Metric && m.Enabled)) continue;
                    current.Issue("CORRECTION_NOT_APPLIED", "Ошибка", "Корректировка не применена: объект / источник не найден либо действие неприменимо. " + c.Reason, c.Metric == "all" ? "" : c.Metric, c.Source, element: c.Element,
                        action: "Проверьте источник, ID, выбранный показатель и действие. Для числовой дельты выберите конкретный показатель.");
                }
                CheckBalances(); AddErrorContours(); current.Corrections = Deserialize<List<Correction>>(Serialize(Config.Corrections.ToList()));
                current.Configuration = Serialize(Config); Last = current;
                DisposeSpatialCalculators();
                if (ViewsEnabled) { progress?.Invoke("Построение проверочных видов..."); Measure("Создание проверочной графики", null, "", () => { CreateGraphics(current); return true; }); }
                RefreshStatuses();
                if (ViewsEnabled && Config.CreateSchedule) { progress?.Invoke("Создание сводной спецификации..."); CreateSchedule(current); }
                RefreshStatuses();
                WriteTimings();
                try { if (!settingsLoadFailed) SaveSettings(); Write("run", current); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { current.Issue("SAVE_FAILED", "Ошибка", "Расчёт выполнен, но запись в RVT не удалась: " + ex.Message); RefreshStatuses(); }
                return current;
            }
            private Detail Row(Record r, Indicator metric, double raw, double factor = 1, bool excluded = false, string reason = "", Solid shape = null, bool volume = false)
            {
                return new Detail
                {
                    Metric = metric.ToString(),
                    SourceKey = r.Source.Key,
                    Source = r.Source.Name,
                    Building = r.Building,
                    Section = r.Section,
                    Level = r.Level?.Name ?? "Без уровня",
                    Elevation = r.Z,
                    Element = IDHelper.ElIdValue(r.Element.Id).ToString(),
                    UniqueId = r.Element.UniqueId,
                    Purpose = r.Role,
                    Apartment = r.Apartment,
                    Profile = r.Profile,
                    Method = current.Method,
                    Unit = Config.Metrics.First(m => m.Key == metric.ToString()).Unit,
                    Raw = raw,
                    Factor = factor,
                    Value = excluded ? 0 : raw * factor,
                    Excluded = excluded,
                    Manual = r.Manual,
                    Reason = reason,
                    Shape = shape,
                    VolumeShape = volume
                };
            }
            private void AddErrorContours()
            {
                if (!ViewsEnabled) return;
                foreach (var issue in current.Issues.Where(x => x.Severity == "Ошибка" && !string.IsNullOrWhiteSpace(x.Element)).GroupBy(x => x.Source + "/" + x.Element).Select(g => g.First()).ToList())
                {
                    try
                    {
                        var source = Sources.FirstOrDefault(s => s.Name == issue.Source && s.Loaded); long id;
                        if (source == null || !long.TryParse(issue.Element, out id)) continue;
                        var element = source.Document.GetElement(IDHelper.CreateElementId(id)); if (element == null) continue;
                        var level = source.Document.GetElement(element.LevelId) as Level;
                        var r = new Record
                        {
                            Source = source,
                            Element = element,
                            Building = string.IsNullOrEmpty(issue.Building) ? "Не определён" : issue.Building,
                            Level = level == null ? null : Config.Levels.FirstOrDefault(l => l.Key == source.Key + "/" + level.UniqueId),
                            Role = "unknown",
                            Profile = Config.Profile
                        };
                        Solid shape = element is SpatialElement ? SpatialPlan(r, "net", "Проверка_ошибок") : null;
                        if (shape == null) foreach (var solid in Solids(element)) shape = Union(shape, Projection(SolidUtils.CreateTransformed(solid, source.Transform)));
                        if (shape == null) continue;
                        current.Details.Add(new Detail
                        {
                            Metric = "Проверка_ошибок",
                            Source = source.Name,
                            SourceKey = source.Key,
                            Building = r.Building,
                            Level = r.Level?.Name ?? "Без уровня",
                            Elevation = r.Z,
                            Element = issue.Element,
                            UniqueId = element.UniqueId,
                            Purpose = "Ошибка",
                            Method = current.Method,
                            Raw = shape.Volume * .09290304,
                            Value = 0,
                            Unit = "м²",
                            Excluded = true,
                            HasError = true,
                            Reason = issue.Code + ": " + issue.Message,
                            Shape = shape
                        });
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch { /* An invalid element can lack even a diagnostic contour; its original structured issue remains. */ }
                }
            }
            private bool VerticalExclusion(Record r, Indicator metric, List<Record> all, Solid shape)
            {
                bool gns = (int)metric <= 4;
                bool stair = r.Role == "stair-gap" && (Dimension(r, "width", true) > 1.5 || Dimension(r, "width", true) > Dimension(r, "stair-width", true));
                bool vertical = r.Role == "multilight" || stair || r.Role == "opening" && (!gns || Config.Method == "new" || shape.Volume * .09290304 > 36) || OneOf(r.Role, "shaft", "engineering-shaft") && (!gns || Config.Method == "new");
                if (!vertical) return false;
                if (string.IsNullOrWhiteSpace(r.Vertical))
                {
                    if (r.Element is Opening)
                    { Notice("VERTICAL_LEVELS", "Ошибка", "Проём задан одним объектом без поэтажных контуров. Для вычитания на верхних этажах задайте зоны с общим ID вертикального пространства.", r, metric.ToString()); return false; }
                    throw new InvalidOperationException("Для многосветного пространства / проёма / шахты нужен общий ID вертикального пространства на всех этажах.");
                }
                var matching = all.Where(x => x.Building == r.Building && x.Section == r.Section && x.Vertical == r.Vertical && x.Level != null && x.Level.Include).ToList();
                if (matching.Count == 0) return false; return r.Z > matching.Min(x => x.Z) + 1e-6;
            }
            private void CalculateAreas(List<Record> records, Indicator metric)
            {
                if ((metric == Indicator.ApartmentsTotal || metric == Indicator.ApartmentsHeated) && string.IsNullOrWhiteSpace(Config.Parameter("apartment")))
                    throw new InvalidOperationException("Для площади квартир задайте параметр ID / номера квартиры.");
                var masks = new Dictionary<string, PlanIndex>(); var used = new Dictionary<string, PlanIndex>();
                int processed = 0;
                var pending = new List<Tuple<Record, Solid, double, bool, string>>();
                foreach (var r in records)
                {
                    Progress("Контуры: " + MetricLabels[metric.ToString()] + " - " + (++processed) + " / " + records.Count + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    // Construction solids are used by envelope/volume routines, never as room floor areas.
                    if (OneOf(r.Role, "structure", "envelope", "footprint", "underground-footprint") && !r.Override.HasValue) continue;
                    if (!(r.Element is SpatialElement) && r.Role == "unknown") continue;
                    if (r.Level == null) { Notice("LEVEL_MISSING", "Ошибка", "Не определён уровень расчётного объекта.", r, metric.ToString()); continue; }
                    if (!r.Level.Include) continue;
                    try
                    {
                        bool include = Eligible(r, metric);
                        var shape = Plan(r, metric); if (shape == null || shape.Volume < 1e-9) continue;
                        if (include && VerticalExclusion(r, metric, records, shape)) include = false;
                        string reason = include ? "Включено по профилю и классификации" : "Исключено по профилю, классификации или этажу";
                        if (r.Override.HasValue) reason = "Ручное правило: " + (include ? "включить" : "исключить");
                        bool reference = r.Source.Mode == "reference";
                        string key = r.Building + "|" + r.Section + "|" + Math.Round(r.Z, 6).ToString("R", CultureInfo.InvariantCulture);
                        if (!include && !reference && (r.Element is SpatialElement || r.Override == false || OneOf(r.Role, "multilight", "stair-gap", "opening", "shaft", "stove", "decoration")))
                        { PlanIndex index; if (!masks.TryGetValue(key, out index)) masks[key] = index = new PlanIndex(); index.Add(shape, r); }
                        pending.Add(Tuple.Create(r, shape, Factor(r, metric), include && !reference, reference ? "Источник только для графической сверки" : reason));
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("AREA_INPUT", "Ошибка", ex.Message, r, metric.ToString()); }
                }
                processed = 0;
                foreach (var item in pending.OrderBy(x => x.Item1.Key, StringComparer.Ordinal))
                {
                    Progress("Пересечения: " + MetricLabels[metric.ToString()] + " - " + (++processed) + " / " + pending.Count + "; ID " + IDHelper.ElIdValue(item.Item1.Element.Id));
                    var r = item.Item1; var shape = item.Item2; string key = r.Building + "|" + r.Section + "|" + Math.Round(r.Z, 6).ToString("R", CultureInfo.InvariantCulture);
                    if (!item.Item4) { current.Details.Add(Row(r, metric, shape.Volume * .09290304, item.Item3, true, item.Item5, shape)); continue; }
                    try
                    {
                        PlanIndex index;
                        if (masks.TryGetValue(key, out index)) foreach (var mask in index.Query(shape))
                            { Progress("Вычитание исключений: " + metric + "; ID " + IDHelper.ElIdValue(r.Element.Id)); shape = Subtract(shape, mask); if (shape == null || shape.Volume < 1e-9) break; }
                        var unique = shape;
                        if (used.TryGetValue(key, out index)) foreach (var previous in index.Query(shape))
                            { Progress("Проверка пересечений: " + metric + "; ID " + IDHelper.ElIdValue(r.Element.Id)); unique = Subtract(unique, previous); if (unique == null || unique.Volume < 1e-9) break; }
                        double duplicate = Math.Max(0, ((shape?.Volume ?? 0) - (unique?.Volume ?? 0)) * .09290304);
                        if (duplicate > .005) Notice("AREA_OVERLAP", "Предупреждение", "Пересечение расчётных контуров " + duplicate.ToString("0.###") + " м² учтено один раз. Проверьте дубликаты и принадлежность квартиры / функции.", r, metric.ToString());
                        if (!used.TryGetValue(key, out index)) used[key] = index = new PlanIndex();
                        index.Add(unique, r);
                        double raw = unique == null ? 0 : unique.Volume * .09290304;
                        if (!IsPublic(r) && r.Role == "transition" && GrossMetric(metric))
                        {
                            var buildings = (Mapped(r.Element, "transition") ?? "").Split(';').Select(x => x.Trim()).Where(x => x.Length > 0).Distinct().ToList();
                            if (buildings.Count < 2) throw new InvalidOperationException("Для перехода укажите минимум два соединяемых корпуса через ;.");
                            foreach (var building in buildings) { var split = r.Copy(); split.Building = building; current.Details.Add(Row(split, metric, raw, 1.0 / buildings.Count, false, "Переход разделён поровну между корпусами", unique)); }
                        }
                        else current.Details.Add(Row(r, metric, raw, item.Item3, false, item.Item5, unique));
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("AREA_BOOLEAN", "Ошибка", ex.Message, r, metric.ToString()); }
                }
            }
            private void CalculateCounts(List<Record> records, Indicator metric)
            {
                bool apartments = metric == Indicator.ApartmentsCount;
                string mode = apartments ? Config.ApartmentMode : Config.ParkingMode;
                if (mode == "id" && string.IsNullOrWhiteSpace(Config.Parameter(apartments ? "apartment" : "parking")))
                    throw new InvalidOperationException("Для подсчёта по ID задайте параметр идентификатора " + (apartments ? "квартиры" : "машино-места") + ".");
                var keys = new HashSet<string>();
                foreach (var r in records.Where(x => x.Source.Mode != "exclude"))
                {
                    if (r.Override == false || Eq(Mapped(r.Element, "include"), "0") || Eq(Mapped(r.Element, "include"), "нет")) continue;
                    if (apartments ? (mode == "instances" ? !(r.Element is FamilyInstance) || (r.Role != "apartment-family" && r.Override != true) : string.IsNullOrWhiteSpace(r.Apartment)) : r.Role != "parking") continue;
                    if (!apartments && mode == "instances" && !(r.Element is FamilyInstance)) continue;
                    try
                    {
                        if (!apartments && Above(r)) continue;
                        if (r.Level == null || !r.Level.Include) continue;
                        string id = apartments ? r.Apartment : Mapped(r.Element, "parking");
                        if (mode == "id" && string.IsNullOrWhiteSpace(id)) throw new InvalidOperationException("Не заполнен идентификатор " + (apartments ? "квартиры" : "машино-места") + ".");
                        if (apartments && string.IsNullOrWhiteSpace(r.Section)) Notice("APARTMENT_SECTION", "Предупреждение", "Секция не задана. Номер квартиры должен быть уникален во всём корпусе.", r, metric.ToString());
                        string key = r.Source.Key + "|" + r.Building + "|" + r.Section + "|" + (mode == "id" ? id : r.Element.UniqueId);
                        bool duplicate = !keys.Add(key); bool excluded = duplicate || r.Source.Mode == "reference";
                        if (duplicate && !apartments) Notice("PARKING_DUPLICATE", "Предупреждение", "Повторный ID машино-места учтён один раз: " + id, r, metric.ToString());
                        Solid shape = null; try { shape = Plan(r, Indicator.PublicRooms); }
                        catch (System.OperationCanceledException) { throw; }
                        catch (Exception ex) { Notice("COUNT_GRAPHICS", "Предупреждение", "Количество определено, но нет контура: " + ex.Message, r, metric.ToString()); }
                        current.Details.Add(Row(r, metric, 1, 1, excluded, duplicate ? "Повторный объект той же квартиры / места" : "Уникальный идентификатор: " + id, shape));
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("COUNT_INPUT", "Ошибка", ex.Message, r, metric.ToString()); }
                }
            }
            private void CalculateFloors(List<Record> records, Indicator metric)
            {
                foreach (var g in records.Where(r => r.Source.Mode == "include" && r.Level != null && (r.Element is SpatialElement || r.Element is Floor || OneOf(r.Role, "envelope", "footprint", "apartment-family"))).GroupBy(r => r.Building + "|" + r.Section + "|" + Math.Round(r.Z, 6)))
                {
                    try
                    {
                        var eligible = g.Where(r => CountLevel(r, metric == Indicator.Storeys)).ToList(); if (eligible.Count == 0) continue;
                        current.Details.Add(Row(eligible[0], metric, 1, 1, false, "Один этаж на корпус и секцию; совпадающие отметки объединены"));
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("FLOOR_INPUT", "Ошибка", ex.Message, g.First(), metric.ToString()); }
                }
            }
            private static void SignaturePart(StringBuilder builder, string value)
            { if (value == null) builder.Append("-1:"); else builder.Append(value.Length).Append(':').Append(value); }
            private string VolumeInputKey(List<Record> records)
            {
                // Per-run cache. Source transforms and document geometry cannot change during the read phase.
                // Include resolved rule/level state, and rebind cached fragments to this metric's records.
                var key = new StringBuilder(); SignaturePart(key, Config.Method); SignaturePart(key, Config.Phase);
                foreach (var r in records.OrderBy(x => x.Key, StringComparer.Ordinal))
                    foreach (var value in new[]{r.Key,r.Source.Mode,r.Building,r.Section,r.Profile,r.BuildingClass,r.Role,r.Part,r.Apartment,r.Vertical,
                        r.Override.HasValue?r.Override.Value.ToString():null,r.Factor?.ToString("R",CultureInfo.InvariantCulture),r.Manual.ToString(),r.SingleStorey.ToString(),r.BuildingIncluded.ToString(),
                        r.Level?.Key,r.Level?.Include.ToString(),r.Level?.Kind,r.Level?.Above,r.Level?.TopSlab,r.Level?.Height,r.Level?.RoofRatio,r.Level?.RoofArea,
                        r.Z.ToString("R",CultureInfo.InvariantCulture)}) SignaturePart(key, value);
                return key.ToString();
            }
            private List<Solid> VolumeSolids(Record r, string metric)
            {
                List<Solid> world; if (worldVolumeSolids.TryGetValue(r.Key, out world)) return world;
                List<Solid> local;
                if (!rawVolumeSolids.TryGetValue(r.Element, out local))
                {
                    Progress("Геометрия конструкций: " + r.Source.Name + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    local = Measure("Геометрия конструкций", r, metric, () => Solids(r.Element)); rawVolumeSolids[r.Element] = local;
                }
                world = local.Select(s => SolidUtils.CreateTransformed(s, r.Source.Transform)).ToList(); worldVolumeSolids[r.Key] = world; return world;
            }
            private Solid RemoveNearby(Solid shape, PlanIndex index, Record r, string metric, string stage)
            {
                foreach (var other in index.Query(shape))
                {
                    Progress(stage + ": " + r.Building + " / " + r.Source.Name + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    try { shape = Measure(stage, r, metric, () => Subtract(shape, other)); }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { throw new InvalidOperationException("Этап: " + stage + ". Объект A: [" + BooleanObject(r) + "]. Объект B: [" + BooleanObject(index.OwnerOf(other)) + "]. " + ex.Message, ex); }
                    if (shape == null || shape.Volume < 1e-9) break;
                }
                return shape;
            }
            private static double CheckedVolume(Solid solid)
            {
                if (solid == null) return 0;
                double value = solid.Volume;
                if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) throw new InvalidOperationException("Недопустимый объём результата геометрической операции: " + value.ToString("R", CultureInfo.InvariantCulture) + " фут³.");
                return value;
            }
            private static double VolumeTolerance(double value) { return Math.Max(1e-7, Math.Abs(value) * 1e-8); }
            private static List<Solid> DifferenceResult(Solid original, Solid difference, Solid intersection = null)
            {
                if (difference == null) throw new InvalidOperationException("Revit не вернул тело результата вычитания.");
                double before = CheckedVolume(original), after = CheckedVolume(difference), tolerance = VolumeTolerance(before);
                if (after > before + tolerance) throw new InvalidOperationException("Вычитание увеличило объём исходного тела.");
                if (intersection != null && Math.Abs(before - CheckedVolume(intersection) - after) > tolerance)
                    throw new InvalidOperationException("Нарушен баланс объёма после вычитания пересечения.");
                return after > 0 ? new List<Solid> { difference } : new List<Solid>();
            }
            private Solid VolumeBoolean(Solid a, Solid b, BooleanOperationsType operation, Record r, string metric, VolumeRetryBudget budget)
            {
                budget?.Spend();
                string stage = operation == BooleanOperationsType.Intersect ? "Проверка фактического пересечения объёмов" : "Вычитание объёмов";
                Progress(stage + ": " + r.Building + " / " + r.Source.Name + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                return Measure(stage, r, metric, () => BooleanOperationsUtils.ExecuteBooleanOperation(a, b, operation));
            }
            private List<Solid> ConnectedVolumeParts(Solid solid, Record r, string metric, VolumeRetryBudget budget)
            {
                budget.Spend(); Progress("Разбиение объёма на связные части: " + r.Building + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                var parts = Measure("Разбиение объёмов на связные части", r, metric, () => SolidUtils.SplitVolumes(solid).ToList());
                double before = CheckedVolume(solid), after = parts.Sum(CheckedVolume);
                if (Math.Abs(before - after) > VolumeTolerance(before) || before > 0 && !parts.Any(p => CheckedVolume(p) > 0))
                    throw new InvalidOperationException("Разбиение на связные части не сохранило объём исходного тела.");
                return parts.Where(p => CheckedVolume(p) > 0).ToList();
            }
            private List<Solid> NormalizedVolumeDifference(Solid a, Solid b, Record r, string metric, VolumeRetryBudget budget)
            {
                var bounds = Bounds(a); var origin = new XYZ((bounds[0] + bounds[3]) / 2, (bounds[1] + bounds[4]) / 2, (bounds[2] + bounds[5]) / 2);
                var errors = new List<string>();
                foreach (double scale in new[] { 1.0, 1000.0 })
                {
                    try
                    {
                        Progress("Вычитание в локальных координатах: " + r.Building + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                        var transform = Transform.Identity; transform.BasisX = XYZ.BasisX * scale; transform.BasisY = XYZ.BasisY * scale; transform.BasisZ = XYZ.BasisZ * scale; transform.Origin = -origin * scale;
                        var left = SolidUtils.CreateTransformed(a, transform); var right = SolidUtils.CreateTransformed(b, transform); double cube = scale * scale * scale;
                        if (Math.Abs(CheckedVolume(left) / cube - CheckedVolume(a)) > VolumeTolerance(CheckedVolume(a)) || Math.Abs(CheckedVolume(right) / cube - CheckedVolume(b)) > VolumeTolerance(CheckedVolume(b)))
                            throw new InvalidOperationException("Преобразование изменило исходный объём.");
                        var difference = VolumeBoolean(left, right, BooleanOperationsType.Difference, r, metric, budget);
                        var intersection = VolumeBoolean(left, right, BooleanOperationsType.Intersect, r, metric, budget);
                        if (intersection == null || CheckedVolume(intersection) > Math.Min(CheckedVolume(left), CheckedVolume(right)) + VolumeTolerance(CheckedVolume(left))) throw new InvalidOperationException("Некорректное контрольное пересечение.");
                        var verified = DifferenceResult(left, difference, intersection); var restored = new List<Solid>();
                        foreach (var part in verified) restored.Add(SolidUtils.CreateTransformed(part, transform.Inverse));
                        if (Math.Abs(restored.Sum(CheckedVolume) + CheckedVolume(intersection) / cube - CheckedVolume(a)) > VolumeTolerance(CheckedVolume(a)))
                            throw new InvalidOperationException("Обратное преобразование нарушило баланс объёмов.");
                        booleanNormalizedRecoveries++; return restored;
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { errors.Add("Масштаб " + scale.ToString(CultureInfo.InvariantCulture) + ": " + ex.Message); }
                }
                throw new InvalidOperationException(string.Join(" | ", errors));
            }
            private List<Solid> SubtractVolume(Solid a, Solid b, Record r, string metric, VolumeRetryBudget budget = null, bool allowSplit = true)
            {
                if (a == null || CheckedVolume(a) == 0) return new List<Solid>();
                if (b == null || CheckedVolume(b) == 0 || !Overlaps(Bounds(a), Bounds(b))) return new List<Solid> { a };
                var failures = new List<string>(); List<Solid> planarResult; string planarFailure;
                if (TryLayeredDifference(a, b, out planarResult, out planarFailure)) return planarResult;
                failures.Add("Плоское разложение: " + planarFailure);
                try { return DifferenceResult(a, VolumeBoolean(a, b, BooleanOperationsType.Difference, r, metric, budget)); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { failures.Add("Прямое вычитание: " + ex.Message); }
                budget = budget ?? new VolumeRetryBudget();
                Solid intersection = null;
                try
                {
                    intersection = VolumeBoolean(a, b, BooleanOperationsType.Intersect, r, metric, budget);
                    if (intersection == null) throw new InvalidOperationException("Revit не вернул результат проверки пересечения.");
                    double overlap = CheckedVolume(intersection);
                    // Only an exactly empty intersection skips subtraction. Never dismiss a small positive overlap.
                    if (overlap == 0) { booleanTouchSkips++; return new List<Solid> { a }; }
                    if (overlap > Math.Min(CheckedVolume(a), CheckedVolume(b)) + VolumeTolerance(CheckedVolume(a)))
                        throw new InvalidOperationException("Пересечение больше одного из исходных тел.");
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { intersection = null; failures.Add("Проверка пересечения: " + ex.Message); }
                if (allowSplit)
                {
                    try
                    {
                        var left = ConnectedVolumeParts(a, r, metric, budget); var right = ConnectedVolumeParts(b, r, metric, budget);
                        if (left.Count > 1 || right.Count > 1)
                        {
                            var recovered = new List<Solid>();
                            foreach (var part in left)
                            {
                                var remaining = new List<Solid> { part };
                                foreach (var cutter in right)
                                {
                                    var next = new List<Solid>();
                                    foreach (var piece in remaining) next.AddRange(SubtractVolume(piece, cutter, r, metric, budget, false));
                                    remaining = next; if (remaining.Count == 0) break;
                                }
                                recovered.AddRange(remaining);
                            }
                            double volume = recovered.Sum(CheckedVolume), before = CheckedVolume(a);
                            if (volume > before + VolumeTolerance(before) || intersection != null && Math.Abs(before - CheckedVolume(intersection) - volume) > VolumeTolerance(before))
                                throw new InvalidOperationException("Нарушен баланс объёма после вычитания связных частей.");
                            booleanSplitRecoveries++; return recovered;
                        }
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { failures.Add("Вычитание связных частей: " + ex.Message); }
                }
                if (intersection != null)
                {
                    try
                    {
                        var recovered = DifferenceResult(a, VolumeBoolean(a, intersection, BooleanOperationsType.Difference, r, metric, budget), intersection);
                        booleanIntersectionRecoveries++; return recovered;
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { failures.Add("Вычитание A ∩ B: " + ex.Message); }
                }
                try { return NormalizedVolumeDifference(a, b, r, metric, budget); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { failures.Add("Локальные координаты: " + ex.Message); }
                throw new InvalidOperationException("Не удалось точно вычесть объёмы после резервных попыток. " + string.Join(" | ", failures));
            }
            private List<Solid> RemoveNearbyVolumes(List<Solid> shapes, PlanIndex index, Record r, string metric, string stage = "Вычитание объёмов")
            {
                var result = new List<Solid>();
                foreach (var original in shapes.Where(s => s != null && CheckedVolume(s) > 0))
                {
                    var remaining = new List<Solid> { original };
                    // All recovered pieces are subsets of original, so its candidate set remains sufficient.
                    foreach (var other in index.Query(original))
                    {
                        var next = new List<Solid>();
                        foreach (var piece in remaining)
                        {
                            try { next.AddRange(SubtractVolume(piece, other, r, metric)); }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { throw new InvalidOperationException("Этап: " + stage + ". Объект A: [" + BooleanObject(r) + "]. Объект B: [" + BooleanObject(index.OwnerOf(other)) + "]. " + ex.Message, ex); }
                        }
                        remaining = next; if (remaining.Count == 0) break;
                    }
                    result.AddRange(remaining);
                }
                return result;
            }
            private VolumeSet BuildVolumeSet(List<Record> records, string metric)
            {
                var result = new VolumeSet(); var inputs = new List<Tuple<Record, Solid>>();
                var ordered = records.OrderBy(r => r.Key, StringComparer.Ordinal).ToList();
                var levels = records.Where(r => r.Level != null && r.Level.Include && r.Level.Kind != "exclude").Select(r => r.Z).ToList();
                if (levels.Count == 0) throw new InvalidOperationException("Не определён нижний расчётный этаж.");
                double lowest = levels.Min();
                var envelopes = ordered.Where(r => r.Role == "envelope" && r.Override != false).ToList();
                result.Reconstructed = envelopes.Count == 0; bool foundSpace = false;
                int processed = 0;
                foreach (var r in envelopes.Count > 0 ? envelopes : ordered)
                {
                    Progress("Подготовка объёмов: " + (++processed) + " / " + (envelopes.Count > 0 ? envelopes.Count : ordered.Count) + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    if (r.Override == false) continue;
                    try
                    {
                        if (envelopes.Count > 0)
                        {
                            var solids = VolumeSolids(r, metric); if (solids.Count == 0) throw new InvalidOperationException("У расчётной оболочки нет замкнутого тела.");
                            foreach (var solid in solids) inputs.Add(Tuple.Create(r, solid));
                        }
                        else if (r.Element is SpatialElement && !(r.Element is Area) && !OneOf(r.Role, "balcony", "terrace", "canopy", "porch", "pit", "passage", "ventilated-void", "soil-filled", "decoration", "french-balcony"))
                        { inputs.Add(Tuple.Create(r, SpatialVolume(r))); foundSpace = true; }
                        else if (r.Role == "structure") foreach (var solid in VolumeSolids(r, metric)) inputs.Add(Tuple.Create(r, solid));
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        // An incomplete explicit envelope cannot stand in for the intended closed building body.
                        if (envelopes.Count > 0) throw new InvalidOperationException("Оболочка " + r.Source.Name + " / ID " + IDHelper.ElIdValue(r.Element.Id) + ": " + ex.Message, ex);
                        result.Problems.Add(Tuple.Create(r.Key, "VOLUME_PART", ex.Message)); Notice("VOLUME_PART", "Ошибка", ex.Message, r, metric);
                    }
                }
                if (result.Reconstructed && !foundSpace) throw new InvalidOperationException("Нет замкнутого внешнего объёма или объёмов Rooms/Spaces. Сумма материалов стен и перекрытий не является строительным объёмом.");
                if (inputs.Count == 0) throw new InvalidOperationException("Расчётная оболочка пуста.");
                var masks = new PlanIndex(true);
                foreach (var r in ordered.Where(r => r.Override == false || OneOf(r.Role, "balcony", "terrace", "canopy", "passage", "decoration", "ventilated-void", "soil-filled")))
                {
                    Progress("Подготовка исключений объёма: " + r.Building + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        if (r.Element is SpatialElement && !(r.Element is Area)) masks.Add(SpatialVolume(r), r);
                        else foreach (var solid in VolumeSolids(r, metric)) masks.Add(solid, r);
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("VOLUME_MASK", "Ошибка", ex.Message, r, metric); throw new InvalidOperationException("Не удалось построить исключаемый объём " + r.Source.Name + " / ID " + IDHelper.ElIdValue(r.Element.Id) + ": " + ex.Message, ex); }
                }
                var expandedInputs = new List<Tuple<Record, Solid>>();
                foreach (var input in inputs) foreach (var part in DecomposeVolumeInput(input.Item2, input.Item1, metric)) expandedInputs.Add(Tuple.Create(input.Item1, part));
                inputs = expandedInputs;
                var used = new PlanIndex(true); processed = 0;
                foreach (var input in inputs)
                {
                    var r = input.Item1;
                    Progress("Объёмные фрагменты: " + (++processed) + " / " + inputs.Count + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        // Clipping, differences and union distribution commute. Never construct a whole-building solid.
                        var shape = Measure("Обрезка по нижнему этажу", r, metric, () => Half(input.Item2, lowest, true));
                        var fragments = RemoveNearbyVolumes(new List<Solid> { shape }, masks, r, metric, "Исключаемые помещения / элементы");
                        fragments = RemoveNearbyVolumes(fragments, used, r, metric, "Устранение повторного учёта пересекающихся тел");
                        // Publish only after every subtraction for this input succeeds; never keep a partial retry.
                        foreach (var fragment in fragments.Where(f => CheckedVolume(f) >= 1e-9))
                        { used.Add(fragment, r); result.Fragments.Add(new VolumeFragment { RecordKey = r.Key, Shape = fragment }); }
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { result.Problems.Add(Tuple.Create(r.Key, "VOLUME_BOOLEAN", ex.Message)); Notice("VOLUME_BOOLEAN", "Ошибка", ex.Message, r, metric); }
                }
                return result;
            }
            private VolumeSet BuildingVolume(List<Record> records, string metric)
            {
                string key = VolumeInputKey(records); VolumeSet result;
                if (volumeSets.TryGetValue(key, out result))
                { volumeCacheHits++; Progress("Повторное использование объёмных фрагментов: " + records[0].Building + " / " + MetricLabel(metric)); }
                else
                {
                    result = Measure("Подготовка объёмных фрагментов", null, metric, () => BuildVolumeSet(records, metric));
                    volumeSets.Add(key, result);
                }
                var lookup = records.ToDictionary(r => r.Key, StringComparer.Ordinal);
                if (result.Reconstructed) Notice("VOLUME_ENVELOPE", "Предупреждение", "Объём восстановлен из пространств и ограждений с однократным учётом пересечений. Проверьте полноту внешней оболочки, в том числе чердаки и надстройки; для подтверждённого результата задайте замкнутую расчётную оболочку.", records[0], metric);
                // Cached partial results must replay their diagnostics for every affected indicator.
                foreach (var problem in result.Problems) Notice(problem.Item2, "Ошибка", problem.Item3, lookup[problem.Item1], metric);
                return result;
            }
            private Solid VolumeSlice(VolumeFragment fragment, Record r, Indicator metric)
            {
                if (metric == Indicator.Volume) return fragment.Shape;
                if (!fragment.Split)
                {
                    Progress("Разделение по нулю здания: " + r.Building + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    Measure("Разделение надземного / подземного объёма", r, metric.ToString(), () =>
                    {
                        fragment.Above = Half(fragment.Shape, zero, true); fragment.Below = Half(fragment.Shape, zero, false);
                        double total = fragment.Shape.Volume, parts = (fragment.Above?.Volume ?? 0) + (fragment.Below?.Volume ?? 0);
                        if (Math.Abs(total - parts) > Math.Max(1e-6, total * 1e-7)) throw new InvalidOperationException("Нарушен баланс надземного и подземного фрагментов после разреза по нулю здания.");
                        fragment.Split = true; return true;
                    });
                }
                return metric == Indicator.VolumeAbove ? fragment.Above : fragment.Below;
            }
            private void CalculateVolumes(List<Record> records, Indicator metric)
            {
                foreach (var group in records.Where(r => r.Source.Mode != "exclude").GroupBy(r => new { r.Building, r.Section, Reference = r.Source.Mode == "reference", ReferenceKey = r.Source.Mode == "reference" ? r.Source.Key : "" }))
                {
                    try
                    {
                        var list = group.ToList(); var volume = BuildingVolume(list, metric.ToString()); var lookup = list.ToDictionary(r => r.Key, StringComparer.Ordinal); int processed = 0;
                        foreach (var fragment in volume.Fragments)
                        {
                            var r = lookup[fragment.RecordKey];
                            Progress("Результат объёма: " + (++processed) + " / " + volume.Fragments.Count + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                            try
                            {
                                var shape = VolumeSlice(fragment, r, metric); if (shape == null || shape.Volume < 1e-9) continue;
                                current.Details.Add(Row(r, metric, shape.Volume * .028316846592, 1, group.Key.Reference, group.Key.Reference ? "Справочная геометрия источника, без включения в итог" : "Фрагмент внешнего объёма исходного объекта; исключения вычтены, пересечения учтены один раз", shape, true));
                            }
                            catch (System.OperationCanceledException) { throw; }
                            catch (Exception ex) { Notice("VOLUME_SLICE", "Ошибка", ex.Message, r, metric.ToString()); }
                        }
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("VOLUME_INPUT", "Ошибка", ex.Message, group.First(), metric.ToString()); }
                }
            }
            private void CalculateFootprints(List<Record> records)
            {
                foreach (var group in records.Where(r => r.Source.Mode != "exclude").GroupBy(r => new { r.Building, r.Section, Reference = r.Source.Mode == "reference", ReferenceKey = r.Source.Mode == "reference" ? r.Source.Key : "" }))
                {
                    var list = group.ToList(); var metric = Indicator.Footprint; var inputs = new List<Tuple<Record, Solid>>();
                    try
                    {
                        var explicitGround = list.Where(r => r.Role == "footprint" && r.Override != false).ToList();
                        if (explicitGround.Count > 0)
                            foreach (var r in explicitGround) { var plan = Plan(r, metric); if (plan != null) inputs.Add(Tuple.Create(r, plan)); }
                        else
                        {
                            var volume = BuildingVolume(list, metric.ToString()); var lookup = list.ToDictionary(r => r.Key, StringComparer.Ordinal); int processed = 0;
                            foreach (var fragment in volume.Fragments)
                            {
                                var r = lookup[fragment.RecordKey];
                                Progress("Проекция застройки: " + (++processed) + " / " + volume.Fragments.Count + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                                Measure("Проекции и сечения для застройки", r, metric.ToString(), () =>
                                {
                                    var plan = Section(fragment.Shape, ground); if (plan != null) inputs.Add(Tuple.Create(r, plan));
                                    var below = Half(fragment.Shape, ground, false);
                                    if (below != null && below.Volume > 1e-9) inputs.Add(Tuple.Create(r, Projection(below)));
                                    return true;
                                });
                            }
                        }
                        foreach (var r in list.Where(r => r.Override != false))
                        {
                            bool residential = !IsPublic(r) || IsHigh(r);
                            bool extra = r.Role == "underground-footprint" || OneOf(r.Role, "porch", "pit", "passage", "terrace", "veranda", "transition") || residential && OneOf(r.Role, "balcony", "loggia", "entrance");
                            if (r.Role == "canopy")
                            {
                                var box = r.Element.get_BoundingBox(null); if (box == null) throw new InvalidOperationException("Не определена высота консоли над землёй.");
                                double lowest = Enumerable.Range(0, 8).Select(i => r.Source.Transform.OfPoint(box.Transform.OfPoint(new XYZ((i & 1) == 0 ? box.Min.X : box.Max.X, (i & 2) == 0 ? box.Min.Y : box.Max.Y, (i & 4) == 0 ? box.Min.Z : box.Max.Z))).Z).Min();
                                extra = lowest - ground < 4.5 / .3048;
                            }
                            if (extra) { var plan = Plan(r, metric); if (plan != null) inputs.Add(Tuple.Create(r, plan)); }
                        }
                        var masks = new PlanIndex(); foreach (var r in list.Where(r => r.Override == false)) masks.Add(Plan(r, metric), r);
                        if (inputs.Count == 0) throw new InvalidOperationException("Пустой контур застройки на отметке земли. Проверьте отметку или задайте контур застройки.");
                        var used = new PlanIndex(); int count = 0;
                        foreach (var input in inputs.OrderBy(x => x.Item1.Key, StringComparer.Ordinal))
                        {
                            var r = input.Item1; Progress("Фрагменты застройки: " + (++count) + " / " + inputs.Count + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                            var shape = RemoveNearby(input.Item2, masks, r, metric.ToString(), "Исключения площади застройки");
                            shape = RemoveNearby(shape, used, r, metric.ToString(), "Пересечения площади застройки");
                            if (shape == null || shape.Volume < 1e-9) continue;
                            var row = Row(r, metric, shape.Volume * .09290304, 1, group.Key.Reference, group.Key.Reference ? "Справочный контур источника, без включения в итог" : "Фрагмент контура застройки исходного объекта; пересечения учтены один раз", shape);
                            row.Level = "План застройки"; row.Elevation = ground; current.Details.Add(row); used.Add(shape, r);
                        }
                        if (explicitGround.Count > 0 && !list.Any(r => r.Role == "underground-footprint"))
                            Notice("FOOTPRINT_UNDERGROUND", "Предупреждение", "Задан ручной наземный контур без подземного. Подтвердите отсутствие выступающих подземных частей либо добавьте их контур.", list.First(), metric.ToString());
                    }
                    catch (System.OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("FOOTPRINT_INPUT", "Ошибка", ex.Message, group.First(), metric.ToString()); }
                }
            }
            private void ApplyDeltas(Indicator metric)
            {
                foreach (var c in Config.Corrections.Where(x => x.Action == "delta" && x.Metric == metric.ToString()))
                {
                    if (string.IsNullOrWhiteSpace(c.Reason)) throw new InvalidOperationException("У корректирующего значения отсутствует обоснование.");
                    double delta = RequiredNumber(c.Value, "корректирующее значение");
                    if ((int)metric >= 21 && delta != Math.Truncate(delta)) throw new InvalidOperationException("Количество и этажность корректируются целыми числами.");
                    string building = string.IsNullOrWhiteSpace(c.Element) ? "Ручная корректировка" : c.Element, section = null;
                    var prior = current.Details.Where(d => d.Metric == metric.ToString() && !d.Excluded).ToList();
                    if (metric == Indicator.Storeys || metric == Indicator.Floors)
                    {
                        var target = (c.Element ?? "").Split('|'); building = target[0]; section = target.Length > 1 ? target[1] : null;
                        var groups = prior.Where(d => Eq(d.Building, building) && (section == null || Eq(d.Section, section))).GroupBy(d => d.Section).ToList();
                        if (groups.Count != 1) throw new InvalidOperationException("Для корректировки этажности укажите существующий корпус и секцию в формате Корпус|Секция.");
                        section = groups[0].Key; prior = groups[0].ToList();
                        if (prior.Sum(d => d.Value) + delta < 0) throw new InvalidOperationException("Корректировка приводит к отрицательному количеству этажей.");
                    }
                    c.Previous = prior.Sum(d => d.Value).ToString("R", CultureInfo.InvariantCulture);
                    current.Details.Add(new Detail
                    {
                        Metric = metric.ToString(),
                        Source = c.Source ?? "Ручной ввод",
                        SourceKey = c.Source,
                        Building = building,
                        Section = section,
                        Level = "Корректировка без контура",
                        Raw = delta,
                        Value = delta,
                        Factor = 1,
                        Manual = true,
                        Unit = Config.Metrics.First(x => x.Key == metric.ToString()).Unit,
                        Method = current.Method,
                        Reason = c.Reason + " | " + c.Author + " | " + c.Date,
                        Profile = Config.Profile
                    });
                    appliedCorrections.Add(c);
                }
            }
            private void Summarize(Metric m)
            {
                var rows = current.Details.Where(d => d.Metric == m.Key && !d.Excluded).ToList();
                bool floors = m.Key == "Storeys" || m.Key == "Floors";
                double value = floors ? rows.GroupBy(d => d.Building + "|" + d.Section).Select(g => g.Sum(d => d.Value)).DefaultIfEmpty(0).Max() : rows.Sum(d => d.Value);
                current.Summary.Add(new Summary
                {
                    Key = m.Key,
                    Name = m.Name,
                    Unit = m.Unit,
                    Value = value,
                    Method = current.Method,
                    Comment = floors ? "Итог проекта - максимальное значение по корпусам / секциям. Поэтажные строки показывают состав." : "",
                    Status = rows.Count == 0 ? "Нет данных" : "Рассчитано"
                });
                if (rows.Count == 0) current.Issue("NO_DATA", "Предупреждение", "Нет включённых объектов показателя; ноль не подтверждает отсутствие таких объектов в проекте.", m.Key);
                if (value < 0) current.Issue("NEGATIVE_TOTAL", "Ошибка", "Отрицательный итог после корректировок. Проверьте знак и величину дельты.", m.Key);
            }
            private void RefreshStatuses()
            {
                foreach (var s in current.Summary)
                {
                    var issues = current.Issues.Where(x => (x.Metric == s.Key || string.IsNullOrEmpty(x.Metric)) && !OneOf(x.Code, "SAVE_FAILED", "SCHEDULE_FAILED", "GRAPHICS_FAILED", "GRAPHIC_ELEMENT", "GRAPHICS_NO_GEOMETRY", "CLEANUP_DEPENDENCY", "REVIT_TRANSACTION")).ToList();
                    if (issues.Any(x => x.Severity == "Ошибка")) s.Status = "Неполный результат";
                    else if (s.Status != "Нет данных" && issues.Any(x => x.Severity == "Предупреждение")) s.Status = "Требует проверки";
                    s.Comment = s.Key == "Storeys" || s.Key == "Floors" ? "Итог проекта - максимум по корпусам и секциям. " : "";
                    s.Comment += string.Join(" | ", issues.Where(x => x.Severity != "Информация").Select(x => x.Code + ": " + x.Message).Distinct().Take(3));
                    s.Value = Math.Abs(s.Value) < 1e-10 ? 0 : s.Value;
                }
            }
            private void CheckBalances()
            {
                var balances = new[] { new[] { "Gns", "GnsResidential", "GnsNonresidential" }, new[] { "GnsResidential", "GnsLivingPart", "GnsNonlivingPart" }, new[] { "Gross", "GrossAbove", "GrossBelow" }, new[] { "Volume", "VolumeAbove", "VolumeBelow" }, new[] { "Np", "NpResidential", "NpNonresidential" } };
                foreach (var keys in balances)
                {
                    var all = keys.Select(k => current.Summary.FirstOrDefault(s => s.Key == k)).ToList(); if (all.Any(x => x == null)) continue;
                    if (Math.Abs(all[0].Value - all[1].Value - all[2].Value) > .01)
                        foreach (var key in keys) current.Issue("BALANCE", "Ошибка", "Не выполнено равенство: " + keys[0] + " = " + keys[1] + " + " + keys[2] + ". Проверьте правила отдельных показателей и классификацию.", key);
                }
                var full = current.Summary.FirstOrDefault(s => s.Key == "ApartmentsTotal"); var heated = current.Summary.FirstOrDefault(s => s.Key == "ApartmentsHeated");
                if (full != null && heated != null && full.Value + .01 < heated.Value) current.Issue("APARTMENT_BALANCE", "Ошибка", "Площадь с летними помещениями меньше площади без них.", "ApartmentsTotal");
            }
            private static Schema StorageSchema()
            {
                var schema = Schema.Lookup(StorageId); if (schema != null) return schema;
                var builder = new SchemaBuilder(StorageId); builder.SetSchemaName("KPLN_TEP_Spec20260826");
                builder.SetReadAccessLevel(AccessLevel.Public); builder.SetWriteAccessLevel(AccessLevel.Public);
                builder.AddSimpleField("Owner", typeof(string)); builder.AddSimpleField("Kind", typeof(string)); builder.AddSimpleField("Payload", typeof(string)); return builder.Finish();
            }
            private static bool Owned(Element e)
            {
                var schema = Schema.Lookup(StorageId); if (schema == null) return false;
                var entity = e.GetEntity(schema); return entity.IsValid() && entity.Get<string>(schema.GetField("Owner")) == Owner;
            }
            private static string Kind(Element e)
            { var s = Schema.Lookup(StorageId); if (s == null) return ""; var entity = e.GetEntity(s); return entity.IsValid() ? entity.Get<string>(s.GetField("Kind")) : ""; }
            private static void Tag(Element e, string kind, string payload = "")
            { var s = StorageSchema(); var entity = new Entity(s); entity.Set(s.GetField("Owner"), Owner); entity.Set(s.GetField("Kind"), kind); entity.Set(s.GetField("Payload"), payload ?? ""); e.SetEntity(entity); }
            private T Read<T>(string kind) where T : class
            {
                var item = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).FirstOrDefault(e => Owned(e) && Kind(e) == kind);
                if (item == null) return null; var schema = StorageSchema();
                try { return Deserialize<T>(UnpackStorage(item.GetEntity(schema).Get<string>(schema.GetField("Payload")))); }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { throw new InvalidOperationException("Не удалось прочитать сохранённые данные ТЭП («" + kind + "»). Данные не перезаписаны. " + ex.Message, ex); }
            }
            private void Write<T>(string kind, T value)
            {
                using (var t = new Transaction(doc, "ТЭП: сохранить " + kind))
                {
                    t.Start(); var data = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).FirstOrDefault(e => Owned(e) && Kind(e) == kind) as DataStorage;
                    if (data == null) { data = DataStorage.Create(doc); data.Name = Owner + "/" + kind; }
                    Tag(data, kind, PackStorage(Serialize(value))); if (t.Commit() != TransactionStatus.Committed) throw new InvalidOperationException("Транзакция записи отменена.");
                }
            }
            private const string CompressedStoragePrefix = "gzip-base64:v1:";
            private static string PackStorage(string text)
            {
                if (text.Length < 1000000) return text;
                using (var output = new MemoryStream())
                {
                    using (var gzip = new System.IO.Compression.GZipStream(output, System.IO.Compression.CompressionMode.Compress, true))
                    { var bytes = Encoding.UTF8.GetBytes(text); gzip.Write(bytes, 0, bytes.Length); }
                    string packed = CompressedStoragePrefix + Convert.ToBase64String(output.ToArray());
                    if (packed.Length >= 16000000) throw new InvalidOperationException("Даже сжатый отчёт превышает лимит одного поля RVT. Экспортируйте отчёт и уменьшите набор показателей.");
                    return packed;
                }
            }
            private static string UnpackStorage(string text)
            {
                if (!text.StartsWith(CompressedStoragePrefix, StringComparison.Ordinal)) return text;
                using (var input = new MemoryStream(Convert.FromBase64String(text.Substring(CompressedStoragePrefix.Length))))
                using (var gzip = new System.IO.Compression.GZipStream(input, System.IO.Compression.CompressionMode.Decompress))
                using (var reader = new StreamReader(gzip, Encoding.UTF8)) return reader.ReadToEnd();
            }
            public static string Serialize<T>(T value)
            { using (var stream = new MemoryStream()) { var serializer = new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings { MaxItemsInObjectGraph = int.MaxValue }); serializer.WriteObject(stream, value); return Encoding.UTF8.GetString(stream.ToArray()); } }
            public static T Deserialize<T>(string text)
            { using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(text))) { return (T)new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings { MaxItemsInObjectGraph = int.MaxValue }).ReadObject(stream); } }
            private sealed class Failures : IFailuresPreprocessor
            {
                private readonly Run run; public Failures(Run value) { run = value; }
                public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
                {
                    bool errors = false; foreach (var failure in accessor.GetFailureMessages())
                    {
                        bool warning = failure.GetSeverity() == FailureSeverity.Warning;
                        run.Issue("REVIT_TRANSACTION", warning ? "Предупреждение" : "Ошибка", failure.GetDescriptionText(), element: string.Join(",", failure.GetFailingElementIds().Select(IDHelper.ElIdValue)));
                        if (warning) accessor.DeleteWarning(failure); else errors = true;
                    }
                    return errors ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
                }
            }
            private void Transaction(string name, Run run, Action action)
            {
                using (var t = new Transaction(doc, name))
                {
                    t.Start(); t.SetFailureHandlingOptions(t.GetFailureHandlingOptions().SetFailuresPreprocessor(new Failures(run)).SetClearAfterRollback(true));
                    action(); if (t.Commit() != TransactionStatus.Committed) throw new InvalidOperationException("Revit отменил транзакцию «" + name + "».");
                }
            }
            private string UniqueName(string desired)
            {
                string safe = new string((desired ?? "ТЭП").Select(c => "\\:{}[]|;<>?`~".Contains(c) ? '_' : c).ToArray()); if (safe.Length > 180) safe = safe.Substring(0, 180);
                var names = new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(View)).Select(v => v.Name)); string name = safe; int i = 2; while (names.Contains(name)) name = safe + " (" + (i++) + ")"; return name;
            }
            private int DeleteOwned(bool graphics, bool schedules, Run run)
            {
                var all = new FilteredElementCollector(doc).WherePasses(new ExtensibleStorageFilter(StorageId)).Where(e => Owned(e) &&
                    (graphics && OneOf(Kind(e), "graphic", "plan", "3d", "graphic-level", "graphic-type") || schedules && OneOf(Kind(e), "schedule", "schedule-key"))).ToList();
                var allowed = new HashSet<long>(all.Select(e => IDHelper.ElIdValue(e.Id))); int count = 0;
                // First delete owned annotations and solids. Only then consider views/levels.
                foreach (var e in all.OrderBy(x => x is View ? 1 : x is Level ? 2 : x is ElementType ? 3 : 0))
                {
                    if (doc.GetElement(e.Id) == null) continue;
                    using (var sub = new SubTransaction(doc))
                    {
                        sub.Start(); var removed = doc.Delete(e.Id);
                        if (removed.Any(id => !allowed.Contains(IDHelper.ElIdValue(id))))
                        { sub.RollBack(); run.Issue("CLEANUP_DEPENDENCY", "Предупреждение", "Служебный объект сохранён: удаление затрагивает не принадлежащие расчёту элементы.", element: IDHelper.ElIdValue(e.Id).ToString(), action: "Проверьте добавленные пользователем аннотации / зависимости."); }
                        else { sub.Commit(); count += removed.Count; }
                    }
                }
                return count;
            }
            public int Cleanup()
            { var report = new Run(); int count = 0; Transaction("ТЭП: удалить служебные результаты", report, () => count = DeleteOwned(true, true, report)); if (Last != null) Last.Issues.AddRange(report.Issues); return count; }
            private static Autodesk.Revit.DB.Color ParseColor(string text)
            {
                var c = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(text); return new Autodesk.Revit.DB.Color(c.R, c.G, c.B);
            }
            private OverrideGraphicSettings Override(Detail d, ElementId fill)
            {
                var colour = ParseColor(d.HasError ? Config.ErrorColor : d.Excluded ? Config.ExcludedColor : d.Manual ? Config.ManualColor :
                    OneOf(d.Purpose, "balcony", "loggia", "terrace", "veranda", "cold-storage") ? Config.SummerColor :
                    OneOf(d.Purpose, "public", "nonresidential-common") ? Config.PublicColor :
                    OneOf(d.Purpose, "technical-room", "technical-space", "technical-void", "roof-vent") ? Config.TechnicalColor : Config.IncludedColor);
                return new OverrideGraphicSettings().SetProjectionLineColor(colour).SetSurfaceForegroundPatternId(fill).SetSurfaceForegroundPatternColor(colour).SetSurfaceTransparency(d.VolumeShape ? 65 : 35);
            }
            public void CreateGraphics(Run run)
            {
                if (!ViewsEnabled) return;
                if (!run.Details.Any(d => d.Shape != null)) { run.Issue("GRAPHICS_NO_GEOMETRY", "Предупреждение", "В сохранённом отчёте нет геометрии Revit. Выполните расчёт заново для построения видов."); return; }
                try
                {
                    Transaction("ТЭП: проверочные виды", run, () =>
                    {
                        if (Config.Graphics == "replace") DeleteOwned(true, false, run);
                        var fill = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(f => f.GetFillPattern().IsSolidFill);
                        if (fill == null) throw new InvalidOperationException("В проекте отсутствует сплошная штриховка.");
                        var regionType = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault();
                        if (regionType == null) throw new InvalidOperationException("В проекте отсутствует тип цветовой области (FilledRegionType).");
                        if (regionType.ForegroundPatternId != fill.Id) { regionType = (FilledRegionType)regionType.Duplicate("ТЭП_Заливка_" + run.Id.Substring(0, 8)); regionType.ForegroundPatternId = fill.Id; Tag(regionType, "graphic-type", run.Id); }
                        var planType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v => v.ViewFamily == ViewFamily.FloorPlan);
                        var threeType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
                        var textType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).FirstOrDefault();
                        if (planType == null) throw new InvalidOperationException("В проекте нет типа плана этажа для проверочных обводок.");
                        var solidsByMetric = new Dictionary<string, List<ElementId>>(); var views3d = new Dictionary<string, View3D>();
                        foreach (var metric in run.Details.Where(d => d.Shape != null).GroupBy(d => d.Metric))
                        {
                            View3D v3 = null; var ids = new List<ElementId>(); solidsByMetric[metric.Key] = ids;
                            if (Config.Create3D && threeType != null)
                            {
                                v3 = View3D.CreateIsometric(doc, threeType.Id); v3.Name = UniqueName(Config.Prefix + metric.Key + "_3D"); Tag(v3, "3d", run.Id); views3d[metric.Key] = v3;
                                var template = doc.GetElement(Config.View3DTemplate ?? "") as View; if (template != null && v3.IsValidViewTemplate(template.Id)) v3.ViewTemplateId = template.Id;
                            }
                            foreach (var floor in metric.GroupBy(d => new { d.Building, d.Section, Z = Math.Round(d.Elevation, 6) }))
                            {
                                ViewPlan plan = null;
                                if (planType != null && floor.Any(d => !d.VolumeShape))
                                {
                                    var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault(l => Math.Abs(l.Elevation - floor.Key.Z) < 1e-6);
                                    if (level == null) { level = Level.Create(doc, floor.Key.Z); Tag(level, "graphic-level", run.Id); }
                                    plan = ViewPlan.Create(doc, planType.Id, level.Id); plan.Name = UniqueName(Config.Prefix + metric.Key + "_" + floor.Key.Building + "_" + floor.First().Level); Tag(plan, "plan", run.Id); plan.Scale = 100;
                                    var template = doc.GetElement(Config.PlanTemplate ?? "") as View; if (template != null && plan.IsValidViewTemplate(template.Id)) plan.ViewTemplateId = template.Id;
                                }
                                bool legend = false;
                                foreach (var detail in floor)
                                {
                                    Progress("Построение видов: " + detail.Metric + " / " + detail.Building + " / " + detail.Level + "; ID " + detail.Element);
                                    using (var sub = new SubTransaction(doc))
                                    {
                                        sub.Start(); try
                                        {
                                            var overrides = Override(detail, fill.Id); var localShapes = new List<Solid>();
                                            if (detail.VolumeShape) localShapes.Add(detail.Shape);
                                            else
                                            {
                                                foreach (PlanarFace face in detail.Shape.Faces.Cast<Face>().OfType<PlanarFace>().Where(f => f.FaceNormal.Z < -.999999))
                                                {
                                                    var loops = face.GetEdgesAsCurveLoops().Select(l => FlatLoop(l, detail.Elevation)).ToList();
                                                    if (plan != null) { var region = FilledRegion.Create(doc, regionType.Id, plan.Id, loops); Tag(region, "graphic", run.Id + "|" + detail.SourceKey + "|" + detail.Element + "|" + detail.Metric); plan.SetElementOverrides(region.Id, overrides); detail.ViewId = plan.UniqueId; }
                                                    localShapes.Add(Extrude(loops, .1));
                                                }
                                            }
                                            if (v3 != null && localShapes.Count > 0)
                                            {
                                                var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel)); ds.ApplicationId = Owner; ds.ApplicationDataId = run.Id + "/" + detail.Metric + "/" + detail.SourceKey + "/" + detail.Element;
                                                ds.SetShape(localShapes.Cast<GeometryObject>().ToList()); Tag(ds, "graphic", run.Id); ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set(detail.Metric + " | " + detail.Building + " | " + detail.Value.ToString("0.###") + " " + detail.Unit); v3.SetElementOverrides(ds.Id, overrides); ids.Add(ds.Id); if (detail.VolumeShape) detail.ViewId = v3.UniqueId;
                                            }
                                            if (plan != null && textType != null)
                                            {
                                                var box = detail.Shape.GetBoundingBox(); var p = box.Transform.OfPoint(box.Min); p = new XYZ(p.X, p.Y, detail.Elevation);
                                                string label = (detail.HasError ? "ОШИБКА " : detail.Excluded ? "ИСКЛЮЧЕНО " : detail.Manual ? "КОРРЕКТИРОВКА " : "") + detail.Element + ": " + detail.Raw.ToString("0.###") + " × " + detail.Factor.ToString("0.###") + " = " + detail.Value.ToString("0.###") + " " + detail.Unit;
                                                var note = TextNote.Create(doc, plan.Id, p, label, textType.Id); Tag(note, "graphic", run.Id);
                                                if (!legend) { CreateLegend(plan, p + new XYZ(0, 12, 0), regionType.Id, textType.Id, fill.Id, run, detail); legend = true; }
                                            }
                                            sub.Commit();
                                        }
                                        catch (System.OperationCanceledException) { throw; }
                                        catch (Exception ex) { sub.RollBack(); run.Issue("GRAPHIC_ELEMENT", "Ошибка", ex.Message, detail.Metric, detail.Source, detail.Building, detail.Element, "Проверьте замкнутость контуров и ограничения вида Revit."); }
                                    }
                                }
                            }
                        }
                        // Hide only this tool's solids belonging to other indicators on its 3D views.
                        foreach (var pair in views3d)
                        { var other = solidsByMetric.Where(p => p.Key != pair.Key).SelectMany(p => p.Value).Where(id => doc.GetElement(id) != null && doc.GetElement(id).CanBeHidden(pair.Value)).ToList(); if (other.Count > 0) pair.Value.HideElements(other); }
                    });
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { run.Issue("GRAPHICS_FAILED", "Ошибка", "Создание графики отменено целиком: " + ex.Message); foreach (var d in run.Details) d.ViewId = null; }
            }
            private void CreateLegend(ViewPlan view, XYZ origin, ElementId regionType, ElementId textType, ElementId fill, Run run, Detail detail)
            {
                var title = TextNote.Create(doc, view.Id, origin + new XYZ(0, 3, 0), "ТЭП | " + run.Method + " | " + run.Date + "\n" + detail.MetricName + " | " + detail.Building + " | " + detail.Level + "\nКонтур: до коэффициента; подпись и итог: после коэффициента.", textType); Tag(title, "graphic", run.Id);
                var entries = new[] { new Detail { Purpose = "heated", Reason = "Жилая / прочая включённая часть" }, new Detail { Purpose = "public", Reason = "Общественная часть" }, new Detail { Purpose = "balcony", Reason = "Летние помещения" }, new Detail { Purpose = "technical-room", Reason = "Технические помещения" }, new Detail { Excluded = true, Reason = "Исключено / справочный контур" }, new Detail { Manual = true, Reason = "Ручное решение" }, new Detail { HasError = true, Reason = "Ошибка исходных данных" } };
                for (int i = 0; i < entries.Length; i++)
                {
                    var p = origin - new XYZ(0, i * 2.2, 0); var points = new[] { p, p + new XYZ(2, 0, 0), p + new XYZ(2, 1, 0), p + new XYZ(0, 1, 0) };
                    var loop = CurveLoop.Create(Enumerable.Range(0, 4).Select(j => (Curve)Line.CreateBound(points[j], points[(j + 1) % 4])).ToList());
                    var swatch = FilledRegion.Create(doc, regionType, view.Id, new List<CurveLoop> { loop }); Tag(swatch, "graphic", run.Id); view.SetElementOverrides(swatch.Id, Override(entries[i], fill));
                    var note = TextNote.Create(doc, view.Id, p + new XYZ(2.7, 1, 0), entries[i].Reason, textType); Tag(note, "graphic", run.Id);
                }
            }
            private void CreateSchedule(Run run)
            {
                if (!ViewsEnabled || !Config.CreateSchedule) return;
                try
                {
                    Transaction("ТЭП: сводная спецификация", run, () =>
                    {
                        if (Config.Graphics == "replace") DeleteOwned(false, true, run);
                        var schedule = ViewSchedule.CreateKeySchedule(doc, new ElementId(BuiltInCategory.OST_GenericModel)); schedule.Name = UniqueName(Config.Prefix + "Сводные показатели"); Tag(schedule, "schedule", run.Id);
                        schedule.KeyScheduleParameterName = UniqueName(Config.Prefix + "Ключ_" + run.Id.Substring(0, 8));
                        var definition = schedule.Definition;
                        var fields = definition.GetSchedulableFields();
                        var comments = fields.FirstOrDefault(f => IDHelper.ElIdValue(f.ParameterId) == (long)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        if (comments == null) throw new InvalidOperationException("Категория не поддерживает поле комментария для сводной спецификации.");
                        var field = definition.AddField(comments); field.ColumnHeading = "Значение | единица | методика | статус"; field.GridColumnWidth = 2.2;
                        definition.GetField(0).ColumnHeading = "Показатель"; definition.GetField(0).GridColumnWidth = 2.6;
                        doc.Regenerate(); var body = schedule.GetTableData().GetSectionData(SectionType.Body);
                        foreach (var s in run.Summary)
                        {
                            var before = new HashSet<long>(new FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType().Select(e => IDHelper.ElIdValue(e.Id)));
                            int row = body.LastRowNumber + 1; if (!body.CanInsertRow(row)) row = body.FirstRowNumber;
                            body.InsertRow(row); doc.Regenerate();
                            var key = new FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType().FirstOrDefault(e => !before.Contains(IDHelper.ElIdValue(e.Id)));
                            if (key == null) throw new InvalidOperationException("Revit не вернул созданную строку ключевой спецификации.");
                            var name = key.get_Parameter(BuiltInParameter.REF_TABLE_ELEM_NAME); var value = key.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            if (name == null || name.IsReadOnly || value == null || value.IsReadOnly) throw new InvalidOperationException("Параметры строки ключевой спецификации недоступны для записи.");
                            name.Set(s.Name); value.Set(Math.Round(s.Value, Config.Decimals, MidpointRounding.AwayFromZero).ToString("N" + Config.Decimals) + " " + s.Unit + " | " + s.Method + " | " + s.Status); Tag(key, "schedule-key", run.Id);
                        }
                    });
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { run.Issue("SCHEDULE_FAILED", "Ошибка", "Сводная спецификация не создана; транзакция отменена: " + ex.Message); }
            }
            public void NavigateRequested()
            {
                var d = RequestedDetail; if (d == null) return;
                try
                {
                    if (!string.IsNullOrWhiteSpace(d.ViewId)) { var view = doc.GetElement(d.ViewId) as View; if (view != null) ui.ActiveView = view; }
                    var source = Sources.FirstOrDefault(s => s.Key == d.SourceKey); ElementId id = source?.RootLink;
                    if (id == null && source?.Document == doc)
                    { var element = string.IsNullOrWhiteSpace(d.UniqueId) ? null : doc.GetElement(d.UniqueId); long numeric; if (element == null && long.TryParse(d.Element, out numeric)) element = doc.GetElement(IDHelper.CreateElementId(numeric)); id = element?.Id; }
                    if (id != null && doc.GetElement(id) != null) { ui.Selection.SetElementIds(new List<ElementId> { id }); ui.ShowElements(id); }
                }
                catch (System.OperationCanceledException) { throw; }
                catch (Exception ex) { TaskDialog.Show("ТЭП: переход к объекту", ex.Message); }
            }
            public static List<Tuple<string, List<object[]>>> ExportTables(Run run, int decimals)
            {
                var tables = new List<Tuple<string, List<object[]>>>();
                var totals = new List<object[]> { new object[] { "Код", "Показатель", "Исходное значение", "Округлено", "Ед.", "Методика", "Статус", "Комментарий", "Дата", "Автор", "Версия правил" } };
                totals.AddRange(run.Summary.Select(s => new object[] { s.Key, s.Name, s.Value, Math.Round(s.Value, decimals, MidpointRounding.AwayFromZero), s.Unit, s.Method, s.Status, s.Comment, run.Date, run.Author, run.Version }));
                tables.Add(Tuple.Create("Итоги", totals));
                foreach (var kind in new[] { "Корпуса", "Связи", "Этажи" })
                {
                    var rows = new List<object[]> { new object[] { "Показатель", "Группа", "Секция", "Этаж", "Отметка, м", "Исходное значение", "Округлено", "Ед.", "Методика" } };
                    var groups = run.Details.Where(d => !d.Excluded).GroupBy(d => new { d.Metric, Group = kind == "Связи" ? d.Source : d.Building, Section = kind == "Связи" ? "" : d.Section, Level = kind == "Этажи" ? d.Level : "", Z = kind == "Этажи" ? Math.Round(d.Elevation, 6) : 0 });
                    foreach (var g in groups)
                    {
                        double value = g.Sum(d => d.Value); if (kind == "Связи" && (g.Key.Metric == "Storeys" || g.Key.Metric == "Floors")) value = g.GroupBy(d => d.Building + "|" + d.Section).Select(x => x.Sum(d => d.Value)).DefaultIfEmpty(0).Max();
                        rows.Add(new object[] { Catalog().FirstOrDefault(m => m.Key == g.Key.Metric)?.Name ?? g.Key.Metric, g.Key.Group, g.Key.Section, g.Key.Level, g.Key.Z * .3048, value, Math.Round(value, decimals, MidpointRounding.AwayFromZero), g.First().Unit, run.Method });
                    }
                    tables.Add(Tuple.Create(kind, rows));
                }
                foreach (var name in new[] { "Детали площадей", "Детали объёмов", "Квартиры", "Машино-места" })
                {
                    var rows = new List<object[]> { new object[] { "Показатель", "Источник", "Ключ источника", "Корпус", "Секция", "Уровень", "Отметка, м", "ElementId", "UniqueId", "Назначение", "ID квартиры", "Профиль", "Методика", "Исходное значение", "Коэффициент", "Значение", "Округлено", "Ед.", "Исключено", "Вручную", "Обоснование", "Проверочный вид" } };
                    var selected = run.Details.Where(d => name == "Квартиры" ? d.Metric.StartsWith("Apartments") : name == "Машино-места" ? d.Metric == "ParkingCount" : name == "Детали объёмов" ? d.Unit == "м³" : d.Unit == "м²");
                    rows.AddRange(selected.Select(d => new object[] { d.MetricName, d.Source, d.SourceKey, d.Building, d.Section, d.Level, d.Elevation * .3048, d.Element, d.UniqueId, d.PurposeName, d.Apartment, Choices("profile").FirstOrDefault(p => p.Key == d.Profile)?.Label ?? d.Profile, d.Method, d.Raw, d.Factor, d.Value, Math.Round(d.Value, decimals, MidpointRounding.AwayFromZero), d.Unit, d.Excluded ? "Да" : "Нет", d.Manual ? "Да" : "Нет", d.Reason, d.ViewId }));
                    tables.Add(Tuple.Create(name, rows));
                }
                var issues = new List<object[]> { new object[] { "Код", "Важность", "Показатель", "Источник", "Корпус", "Элемент", "Сообщение", "Что сделать" } };
                issues.AddRange(run.Issues.Select(x => new object[] { x.Code, x.Severity, x.MetricName, x.Source, x.Building, x.Element, x.Message, x.Action })); tables.Add(Tuple.Create("Ошибки", issues));
                var settings = new List<object[]> { new object[] { "Раздел", "Значение" }, new object[] { "Методика", run.Method }, new object[] { "Версия правил", run.Version }, new object[] { "Дата", run.Date }, new object[] { "Автор", run.Author } };
                string config = run.Configuration ?? ""; for (int i = 0; i < config.Length; i += 30000) settings.Add(new object[] { "Настройки JSON, часть " + (i / 30000 + 1), config.Substring(i, Math.Min(30000, config.Length - i)) });
                tables.Add(Tuple.Create("Методика и настройки", settings));
                var corrections = new List<object[]> { new object[] { "Показатель", "Действие", "Источник", "Объект / корпус для дельты", "Было", "Новое значение / дельта", "Причина", "Автор", "Дата" } };
                corrections.AddRange(run.Corrections.Select(c => new object[] { c.Metric, c.Action, c.Source, c.Element, c.Previous, c.Value, c.Reason, c.Author, c.Date })); tables.Add(Tuple.Create("Ручные корректировки", corrections));
                return tables;
            }
            public static void ExportRun(Run run, string path, int decimals)
            {
                var tables = ExportTables(run, decimals);
                // Build to a sibling temporary file first: an interrupted export must not destroy an existing report.
                string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) WriteXlsx(temporary, tables);
                    else
                    {
                        using (var writer = new StreamWriter(temporary, false, new UTF8Encoding(true))) foreach (var table in tables)
                            { writer.WriteLine(CsvCell("РАЗДЕЛ: " + table.Item1)); foreach (var row in table.Item2) writer.WriteLine(string.Join(";", row.Select(CsvCell))); writer.WriteLine(); }
                    }
                    if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
            public static string CsvCell(object value)
            {
                string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
                if (value is string && text.TrimStart().Length > 0 && "=+-@".Contains(text.TrimStart()[0])) text = "'" + text;
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }
            private static string ColumnName(int index)
            { string text = ""; for (int n = index + 1; n > 0; n = (n - 1) / 26) text = (char)('A' + (n - 1) % 26) + text; return text; }
            private static string XmlText(object value)
            { var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""; return new string(text.Where(c => XmlConvert.IsXmlChar(c)).ToArray()); }
            private static void WriteXlsx(string path, List<Tuple<string, List<object[]>>> tables)
            {
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                using (var package = Package.Open(path, FileMode.Create, FileAccess.ReadWrite))
                {
                    var workbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                    var workbook = package.CreatePart(workbookUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", CompressionOption.Normal);
                    package.CreateRelationship(workbookUri, TargetMode.Internal, rel.NamespaceName + "/officeDocument");
                    var sheets = new XElement(ns + "sheets"); int sheetIndex = 0;
                    foreach (var table in tables)
                    {
                        sheetIndex++; if (table.Item2.Count > 1048576) throw new InvalidOperationException("Лист превышает лимит Excel. Экспортируйте CSV.");
                        var uri = new Uri("/xl/worksheets/sheet" + sheetIndex + ".xml", UriKind.Relative);
                        var sheet = package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
                        string rid = "rId" + sheetIndex; workbook.CreateRelationship(PackUriHelper.GetRelativeUri(workbookUri, uri), TargetMode.Internal, rel.NamespaceName + "/worksheet", rid);
                        sheets.Add(new XElement(ns + "sheet", new XAttribute("name", table.Item1), new XAttribute("sheetId", sheetIndex), new XAttribute(rel + "id", rid)));
                        var data = new XElement(ns + "sheetData"); int rowIndex = 0;
                        foreach (var values in table.Item2)
                        {
                            rowIndex++; var row = new XElement(ns + "row", new XAttribute("r", rowIndex)); int col = 0;
                            foreach (var value in values)
                            {
                                var cell = new XElement(ns + "c", new XAttribute("r", ColumnName(col++) + rowIndex));
                                if (value is double || value is int || value is long || value is decimal) cell.Add(new XElement(ns + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
                                else
                                {
                                    var text = XmlText(value); if (text.Length > 32767) throw new InvalidOperationException("Текст ячейки превышает лимит Excel. Экспортируйте CSV.");
                                    cell.Add(new XAttribute("t", "inlineStr"), new XElement(ns + "is", new XElement(ns + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
                                }
                                row.Add(cell);
                            }
                            data.Add(row);
                        }
                        var root = new XElement(ns + "worksheet", new XElement(ns + "sheetViews", new XElement(ns + "sheetView", new XAttribute("workbookViewId", 0), new XElement(ns + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")))),
                            new XElement(ns + "cols", new XElement(ns + "col", new XAttribute("min", 1), new XAttribute("max", table.Item2[0].Length), new XAttribute("width", 24), new XAttribute("customWidth", 1))), data,
                            new XElement(ns + "autoFilter", new XAttribute("ref", "A1:" + ColumnName(table.Item2[0].Length - 1) + Math.Max(1, rowIndex))));
                        using (var stream = sheet.GetStream()) new XDocument(root).Save(stream);
                    }
                    using (var stream = workbook.GetStream()) new XDocument(new XElement(ns + "workbook", new XAttribute(XNamespace.Xmlns + "r", rel), sheets)).Save(stream);
                }
            }
        }
    }
}




// Embedded Clipper 6.4.2 for C# 7.3 / Revit 2020 compatibility; no external DLL or extra project file.
// Upstream: https://sourceforge.net/projects/polyclipping/ (6.4.2).
// Source mirror: https://github.com/Geri-Borbas/Clipper/blob/master/C%23/clipper_library/clipper.cs
// Only namespace/usings/preprocessor selection changed.
/*
Boost Software License - Version 1.0 - August 17th, 2003
http://www.boost.org/LICENSE_1_0.txt

Permission is hereby granted, free of charge, to any person or organization
obtaining a copy of the software and accompanying documentation covered by
this license (the "Software") to use, reproduce, display, distribute,
execute, and transmit the Software, and to prepare derivative works of the
Software, and to permit third-parties to whom the Software is furnished to
do so, all subject to the following:

The copyright notices in the Software and this entire statement, including
the above license grant, this restriction and the following disclaimer,
must be included in all copies of the Software, in whole or in part, and
all derivative works of the Software, unless such copies or derivative
works are solely in the form of machine-executable object code generated by
a source language processor.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT
SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE
FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
*/
/*******************************************************************************
*                                                                              *
* Author    :  Angus Johnson                                                   *
* Version   :  6.4.2                                                           *
* Date      :  27 February 2017                                                *
* Website   :  http://www.angusj.com                                           *
* Copyright :  Angus Johnson 2010-2017                                         *
*                                                                              *
* License:                                                                     *
* Use, modification & distribution is subject to Boost Software License Ver 1. *
* http://www.boost.org/LICENSE_1_0.txt                                         *
*                                                                              *
* Attributions:                                                                *
* The code in this library is an extension of Bala Vatti's clipping algorithm: *
* "A generic solution to polygon clipping"                                     *
* Communications of the ACM, Vol 35, Issue 7 (July 1992) pp 56-63.             *
* http://portal.acm.org/citation.cfm?id=129906                                 *
*                                                                              *
* Computer graphics and geometric modeling: implementation and algorithms      *
* By Max K. Agoston                                                            *
* Springer; 1 edition (January 4, 2005)                                        *
* http://books.google.com/books?q=vatti+clipping+agoston                       *
*                                                                              *
* See also:                                                                    *
* "Polygon Offsetting by Computing Winding Numbers"                            *
* Paper no. DETC2005-85513 pp. 565-575                                         *
* ASME 2005 International Design Engineering Technical Conferences             *
* and Computers and Information in Engineering Conference (IDETC/CIE2005)      *
* September 24-28, 2005 , Long Beach, California, USA                          *
* http://www.me.berkeley.edu/~mcmains/pubs/DAC05OffsetPolygon.pdf              *
*                                                                              *
*******************************************************************************/

/*******************************************************************************
*                                                                              *
* This is a translation of the Delphi Clipper library and the naming style     *
* used has retained a Delphi flavour.                                          *
*                                                                              *
*******************************************************************************/

//use_int32: When enabled 32bit ints are used instead of 64bit ints. This
//improve performance but coordinate values are limited to the range +/- 46340
//#define use_int32

//use_xyz: adds a Z member to IntPoint. Adds a minor cost to performance.
//#define use_xyz

//use_lines: Enables open path clipping. Adds a very minor cost to performance.



//using System.Text;          //for Int128.AsString() & StringBuilder
//using System.IO;            //debugging with streamReader & StreamWriter
//using System.Windows.Forms; //debugging to clipboard

namespace KPLN_Tools.ExternalCommands.TepClipper
{
    using System;
    using System.Collections.Generic;

    using cInt = Int64;

    using Path = List<IntPoint>;
    using Paths = List<List<IntPoint>>;

    public struct DoublePoint
    {
        public double X;
        public double Y;

        public DoublePoint(double x = 0, double y = 0)
        {
            this.X = x; this.Y = y;
        }
        public DoublePoint(DoublePoint dp)
        {
            this.X = dp.X; this.Y = dp.Y;
        }
        public DoublePoint(IntPoint ip)
        {
            this.X = ip.X; this.Y = ip.Y;
        }
    };


    //------------------------------------------------------------------------------
    // PolyTree & PolyNode classes
    //------------------------------------------------------------------------------

    public class PolyTree : PolyNode
    {
        internal List<PolyNode> m_AllPolys = new List<PolyNode>();

        //The GC probably handles this cleanup more efficiently ...
        //~PolyTree(){Clear();}

        public void Clear()
        {
            for (int i = 0; i < m_AllPolys.Count; i++)
                m_AllPolys[i] = null;
            m_AllPolys.Clear();
            m_Childs.Clear();
        }

        public PolyNode GetFirst()
        {
            if (m_Childs.Count > 0)
                return m_Childs[0];
            else
                return null;
        }

        public int Total
        {
            get
            {
                int result = m_AllPolys.Count;
                //with negative offsets, ignore the hidden outer polygon ...
                if (result > 0 && m_Childs[0] != m_AllPolys[0]) result--;
                return result;
            }
        }

    }

    public class PolyNode
    {
        internal PolyNode m_Parent;
        internal Path m_polygon = new Path();
        internal int m_Index;
        internal JoinType m_jointype;
        internal EndType m_endtype;
        internal List<PolyNode> m_Childs = new List<PolyNode>();

        private bool IsHoleNode()
        {
            bool result = true;
            PolyNode node = m_Parent;
            while (node != null)
            {
                result = !result;
                node = node.m_Parent;
            }
            return result;
        }

        public int ChildCount
        {
            get { return m_Childs.Count; }
        }

        public Path Contour
        {
            get { return m_polygon; }
        }

        internal void AddChild(PolyNode Child)
        {
            int cnt = m_Childs.Count;
            m_Childs.Add(Child);
            Child.m_Parent = this;
            Child.m_Index = cnt;
        }

        public PolyNode GetNext()
        {
            if (m_Childs.Count > 0)
                return m_Childs[0];
            else
                return GetNextSiblingUp();
        }

        internal PolyNode GetNextSiblingUp()
        {
            if (m_Parent == null)
                return null;
            else if (m_Index == m_Parent.m_Childs.Count - 1)
                return m_Parent.GetNextSiblingUp();
            else
                return m_Parent.m_Childs[m_Index + 1];
        }

        public List<PolyNode> Childs
        {
            get { return m_Childs; }
        }

        public PolyNode Parent
        {
            get { return m_Parent; }
        }

        public bool IsHole
        {
            get { return IsHoleNode(); }
        }

        public bool IsOpen { get; set; }
    }


    //------------------------------------------------------------------------------
    // Int128 struct (enables safe math on signed 64bit integers)
    // eg Int128 val1((Int64)9223372036854775807); //ie 2^63 -1
    //    Int128 val2((Int64)9223372036854775807);
    //    Int128 val3 = val1 * val2;
    //    val3.ToString => "85070591730234615847396907784232501249" (8.5e+37)
    //------------------------------------------------------------------------------

    internal struct Int128
    {
        private Int64 hi;
        private UInt64 lo;

        public Int128(Int64 _lo)
        {
            lo = (UInt64)_lo;
            if (_lo < 0) hi = -1;
            else hi = 0;
        }

        public Int128(Int64 _hi, UInt64 _lo)
        {
            lo = _lo;
            hi = _hi;
        }

        public Int128(Int128 val)
        {
            hi = val.hi;
            lo = val.lo;
        }

        public bool IsNegative()
        {
            return hi < 0;
        }

        public static bool operator ==(Int128 val1, Int128 val2)
        {
            if ((object)val1 == (object)val2) return true;
            else if ((object)val1 == null || (object)val2 == null) return false;
            return (val1.hi == val2.hi && val1.lo == val2.lo);
        }

        public static bool operator !=(Int128 val1, Int128 val2)
        {
            return !(val1 == val2);
        }

        public override bool Equals(System.Object obj)
        {
            if (obj == null || !(obj is Int128))
                return false;
            Int128 i128 = (Int128)obj;
            return (i128.hi == hi && i128.lo == lo);
        }

        public override int GetHashCode()
        {
            return hi.GetHashCode() ^ lo.GetHashCode();
        }

        public static bool operator >(Int128 val1, Int128 val2)
        {
            if (val1.hi != val2.hi)
                return val1.hi > val2.hi;
            else
                return val1.lo > val2.lo;
        }

        public static bool operator <(Int128 val1, Int128 val2)
        {
            if (val1.hi != val2.hi)
                return val1.hi < val2.hi;
            else
                return val1.lo < val2.lo;
        }

        public static Int128 operator +(Int128 lhs, Int128 rhs)
        {
            lhs.hi += rhs.hi;
            lhs.lo += rhs.lo;
            if (lhs.lo < rhs.lo) lhs.hi++;
            return lhs;
        }

        public static Int128 operator -(Int128 lhs, Int128 rhs)
        {
            return lhs + -rhs;
        }

        public static Int128 operator -(Int128 val)
        {
            if (val.lo == 0)
                return new Int128(-val.hi, 0);
            else
                return new Int128(~val.hi, ~val.lo + 1);
        }

        public static explicit operator double(Int128 val)
        {
            const double shift64 = 18446744073709551616.0; //2^64
            if (val.hi < 0)
            {
                if (val.lo == 0)
                    return (double)val.hi * shift64;
                else
                    return -(double)(~val.lo + ~val.hi * shift64);
            }
            else
                return (double)(val.lo + val.hi * shift64);
        }

        //nb: Constructing two new Int128 objects every time we want to multiply longs  
        //is slow. So, although calling the Int128Mul method doesn't look as clean, the 
        //code runs significantly faster than if we'd used the * operator.

        public static Int128 Int128Mul(Int64 lhs, Int64 rhs)
        {
            bool negate = (lhs < 0) != (rhs < 0);
            if (lhs < 0) lhs = -lhs;
            if (rhs < 0) rhs = -rhs;
            UInt64 int1Hi = (UInt64)lhs >> 32;
            UInt64 int1Lo = (UInt64)lhs & 0xFFFFFFFF;
            UInt64 int2Hi = (UInt64)rhs >> 32;
            UInt64 int2Lo = (UInt64)rhs & 0xFFFFFFFF;

            //nb: see comments in clipper.pas
            UInt64 a = int1Hi * int2Hi;
            UInt64 b = int1Lo * int2Lo;
            UInt64 c = int1Hi * int2Lo + int1Lo * int2Hi;

            UInt64 lo;
            Int64 hi;
            hi = (Int64)(a + (c >> 32));

            unchecked { lo = (c << 32) + b; }
            if (lo < b) hi++;
            Int128 result = new Int128(hi, lo);
            return negate ? -result : result;
        }

    };

    //------------------------------------------------------------------------------
    //------------------------------------------------------------------------------

    public struct IntPoint
    {
        public cInt X;
        public cInt Y;
        public IntPoint(cInt X, cInt Y)
        {
            this.X = X; this.Y = Y;
        }
        public IntPoint(double x, double y)
        {
            this.X = (cInt)x; this.Y = (cInt)y;
        }

        public IntPoint(IntPoint pt)
        {
            this.X = pt.X; this.Y = pt.Y;
        }

        public static bool operator ==(IntPoint a, IntPoint b)
        {
            return a.X == b.X && a.Y == b.Y;
        }

        public static bool operator !=(IntPoint a, IntPoint b)
        {
            return a.X != b.X || a.Y != b.Y;
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            if (obj is IntPoint)
            {
                IntPoint a = (IntPoint)obj;
                return (X == a.X) && (Y == a.Y);
            }
            else return false;
        }

        public override int GetHashCode()
        {
            //simply prevents a compiler warning
            return base.GetHashCode();
        }

    }// end struct IntPoint

    public struct IntRect
    {
        public cInt left;
        public cInt top;
        public cInt right;
        public cInt bottom;

        public IntRect(cInt l, cInt t, cInt r, cInt b)
        {
            this.left = l; this.top = t;
            this.right = r; this.bottom = b;
        }
        public IntRect(IntRect ir)
        {
            this.left = ir.left; this.top = ir.top;
            this.right = ir.right; this.bottom = ir.bottom;
        }
    }

    public enum ClipType { ctIntersection, ctUnion, ctDifference, ctXor };
    public enum PolyType { ptSubject, ptClip };

    //By far the most widely used winding rules for polygon filling are
    //EvenOdd & NonZero (GDI, GDI+, XLib, OpenGL, Cairo, AGG, Quartz, SVG, Gr32)
    //Others rules include Positive, Negative and ABS_GTR_EQ_TWO (only in OpenGL)
    //see http://glprogramming.com/red/chapter11.html
    public enum PolyFillType { pftEvenOdd, pftNonZero, pftPositive, pftNegative };

    public enum JoinType { jtSquare, jtRound, jtMiter };
    public enum EndType { etClosedPolygon, etClosedLine, etOpenButt, etOpenSquare, etOpenRound };

    internal enum EdgeSide { esLeft, esRight };
    internal enum Direction { dRightToLeft, dLeftToRight };

    internal class TEdge
    {
        internal IntPoint Bot;
        internal IntPoint Curr; //current (updated for every new scanbeam)
        internal IntPoint Top;
        internal IntPoint Delta;
        internal double Dx;
        internal PolyType PolyTyp;
        internal EdgeSide Side; //side only refers to current side of solution poly
        internal int WindDelta; //1 or -1 depending on winding direction
        internal int WindCnt;
        internal int WindCnt2; //winding count of the opposite polytype
        internal int OutIdx;
        internal TEdge Next;
        internal TEdge Prev;
        internal TEdge NextInLML;
        internal TEdge NextInAEL;
        internal TEdge PrevInAEL;
        internal TEdge NextInSEL;
        internal TEdge PrevInSEL;
    };

    public class IntersectNode
    {
        internal TEdge Edge1;
        internal TEdge Edge2;
        internal IntPoint Pt;
    };

    public class MyIntersectNodeSort : IComparer<IntersectNode>
    {
        public int Compare(IntersectNode node1, IntersectNode node2)
        {
            cInt i = node2.Pt.Y - node1.Pt.Y;
            if (i > 0) return 1;
            else if (i < 0) return -1;
            else return 0;
        }
    }

    internal class LocalMinima
    {
        internal cInt Y;
        internal TEdge LeftBound;
        internal TEdge RightBound;
        internal LocalMinima Next;
    };

    internal class Scanbeam
    {
        internal cInt Y;
        internal Scanbeam Next;
    };

    internal class Maxima
    {
        internal cInt X;
        internal Maxima Next;
        internal Maxima Prev;
    };

    //OutRec: contains a path in the clipping solution. Edges in the AEL will
    //carry a pointer to an OutRec when they are part of the clipping solution.
    internal class OutRec
    {
        internal int Idx;
        internal bool IsHole;
        internal bool IsOpen;
        internal OutRec FirstLeft; //see comments in clipper.pas
        internal OutPt Pts;
        internal OutPt BottomPt;
        internal PolyNode PolyNode;
    };

    internal class OutPt
    {
        internal int Idx;
        internal IntPoint Pt;
        internal OutPt Next;
        internal OutPt Prev;
    };

    internal class Join
    {
        internal OutPt OutPt1;
        internal OutPt OutPt2;
        internal IntPoint OffPt;
    };

    public class ClipperBase
    {
        internal const double horizontal = -3.4E+38;
        internal const int Skip = -2;
        internal const int Unassigned = -1;
        internal const double tolerance = 1.0E-20;
        internal static bool near_zero(double val) { return (val > -tolerance) && (val < tolerance); }

        public const cInt loRange = 0x3FFFFFFF;
        public const cInt hiRange = 0x3FFFFFFFFFFFFFFFL;

        internal LocalMinima m_MinimaList;
        internal LocalMinima m_CurrentLM;
        internal List<List<TEdge>> m_edges = new List<List<TEdge>>();
        internal Scanbeam m_Scanbeam;
        internal List<OutRec> m_PolyOuts;
        internal TEdge m_ActiveEdges;
        internal bool m_UseFullRange;
        internal bool m_HasOpenPaths;

        //------------------------------------------------------------------------------

        public bool PreserveCollinear
        {
            get;
            set;
        }
        //------------------------------------------------------------------------------

        public void Swap(ref cInt val1, ref cInt val2)
        {
            cInt tmp = val1;
            val1 = val2;
            val2 = tmp;
        }
        //------------------------------------------------------------------------------

        internal static bool IsHorizontal(TEdge e)
        {
            return e.Delta.Y == 0;
        }
        //------------------------------------------------------------------------------

        internal bool PointIsVertex(IntPoint pt, OutPt pp)
        {
            OutPt pp2 = pp;
            do
            {
                if (pp2.Pt == pt) return true;
                pp2 = pp2.Next;
            }
            while (pp2 != pp);
            return false;
        }
        //------------------------------------------------------------------------------

        internal bool PointOnLineSegment(IntPoint pt,
            IntPoint linePt1, IntPoint linePt2, bool UseFullRange)
        {
            if (UseFullRange)
                return ((pt.X == linePt1.X) && (pt.Y == linePt1.Y)) ||
                  ((pt.X == linePt2.X) && (pt.Y == linePt2.Y)) ||
                  (((pt.X > linePt1.X) == (pt.X < linePt2.X)) &&
                  ((pt.Y > linePt1.Y) == (pt.Y < linePt2.Y)) &&
                  ((Int128.Int128Mul((pt.X - linePt1.X), (linePt2.Y - linePt1.Y)) ==
                  Int128.Int128Mul((linePt2.X - linePt1.X), (pt.Y - linePt1.Y)))));
            else
                return ((pt.X == linePt1.X) && (pt.Y == linePt1.Y)) ||
                  ((pt.X == linePt2.X) && (pt.Y == linePt2.Y)) ||
                  (((pt.X > linePt1.X) == (pt.X < linePt2.X)) &&
                  ((pt.Y > linePt1.Y) == (pt.Y < linePt2.Y)) &&
                  ((pt.X - linePt1.X) * (linePt2.Y - linePt1.Y) ==
                    (linePt2.X - linePt1.X) * (pt.Y - linePt1.Y)));
        }
        //------------------------------------------------------------------------------

        internal bool PointOnPolygon(IntPoint pt, OutPt pp, bool UseFullRange)
        {
            OutPt pp2 = pp;
            while (true)
            {
                if (PointOnLineSegment(pt, pp2.Pt, pp2.Next.Pt, UseFullRange))
                    return true;
                pp2 = pp2.Next;
                if (pp2 == pp) break;
            }
            return false;
        }
        //------------------------------------------------------------------------------

        internal static bool SlopesEqual(TEdge e1, TEdge e2, bool UseFullRange)
        {
            if (UseFullRange)
                return Int128.Int128Mul(e1.Delta.Y, e2.Delta.X) ==
                    Int128.Int128Mul(e1.Delta.X, e2.Delta.Y);
            else return (cInt)(e1.Delta.Y) * (e2.Delta.X) ==
              (cInt)(e1.Delta.X) * (e2.Delta.Y);
        }
        //------------------------------------------------------------------------------

        internal static bool SlopesEqual(IntPoint pt1, IntPoint pt2,
            IntPoint pt3, bool UseFullRange)
        {
            if (UseFullRange)
                return Int128.Int128Mul(pt1.Y - pt2.Y, pt2.X - pt3.X) ==
                  Int128.Int128Mul(pt1.X - pt2.X, pt2.Y - pt3.Y);
            else return
              (cInt)(pt1.Y - pt2.Y) * (pt2.X - pt3.X) - (cInt)(pt1.X - pt2.X) * (pt2.Y - pt3.Y) == 0;
        }
        //------------------------------------------------------------------------------

        internal static bool SlopesEqual(IntPoint pt1, IntPoint pt2,
            IntPoint pt3, IntPoint pt4, bool UseFullRange)
        {
            if (UseFullRange)
                return Int128.Int128Mul(pt1.Y - pt2.Y, pt3.X - pt4.X) ==
                  Int128.Int128Mul(pt1.X - pt2.X, pt3.Y - pt4.Y);
            else return
              (cInt)(pt1.Y - pt2.Y) * (pt3.X - pt4.X) - (cInt)(pt1.X - pt2.X) * (pt3.Y - pt4.Y) == 0;
        }
        //------------------------------------------------------------------------------

        internal ClipperBase() //constructor (nb: no external instantiation)
        {
            m_MinimaList = null;
            m_CurrentLM = null;
            m_UseFullRange = false;
            m_HasOpenPaths = false;
        }
        //------------------------------------------------------------------------------

        public virtual void Clear()
        {
            DisposeLocalMinimaList();
            for (int i = 0; i < m_edges.Count; ++i)
            {
                for (int j = 0; j < m_edges[i].Count; ++j) m_edges[i][j] = null;
                m_edges[i].Clear();
            }
            m_edges.Clear();
            m_UseFullRange = false;
            m_HasOpenPaths = false;
        }
        //------------------------------------------------------------------------------

        private void DisposeLocalMinimaList()
        {
            while (m_MinimaList != null)
            {
                LocalMinima tmpLm = m_MinimaList.Next;
                m_MinimaList = null;
                m_MinimaList = tmpLm;
            }
            m_CurrentLM = null;
        }
        //------------------------------------------------------------------------------

        void RangeTest(IntPoint Pt, ref bool useFullRange)
        {
            if (useFullRange)
            {
                if (Pt.X > hiRange || Pt.Y > hiRange || -Pt.X > hiRange || -Pt.Y > hiRange)
                    throw new ClipperException("Coordinate outside allowed range");
            }
            else if (Pt.X > loRange || Pt.Y > loRange || -Pt.X > loRange || -Pt.Y > loRange)
            {
                useFullRange = true;
                RangeTest(Pt, ref useFullRange);
            }
        }
        //------------------------------------------------------------------------------

        private void InitEdge(TEdge e, TEdge eNext,
          TEdge ePrev, IntPoint pt)
        {
            e.Next = eNext;
            e.Prev = ePrev;
            e.Curr = pt;
            e.OutIdx = Unassigned;
        }
        //------------------------------------------------------------------------------

        private void InitEdge2(TEdge e, PolyType polyType)
        {
            if (e.Curr.Y >= e.Next.Curr.Y)
            {
                e.Bot = e.Curr;
                e.Top = e.Next.Curr;
            }
            else
            {
                e.Top = e.Curr;
                e.Bot = e.Next.Curr;
            }
            SetDx(e);
            e.PolyTyp = polyType;
        }
        //------------------------------------------------------------------------------

        private TEdge FindNextLocMin(TEdge E)
        {
            TEdge E2;
            for (; ; )
            {
                while (E.Bot != E.Prev.Bot || E.Curr == E.Top) E = E.Next;
                if (E.Dx != horizontal && E.Prev.Dx != horizontal) break;
                while (E.Prev.Dx == horizontal) E = E.Prev;
                E2 = E;
                while (E.Dx == horizontal) E = E.Next;
                if (E.Top.Y == E.Prev.Bot.Y) continue; //ie just an intermediate horz.
                if (E2.Prev.Bot.X < E.Bot.X) E = E2;
                break;
            }
            return E;
        }
        //------------------------------------------------------------------------------

        private TEdge ProcessBound(TEdge E, bool LeftBoundIsForward)
        {
            TEdge EStart, Result = E;
            TEdge Horz;

            if (Result.OutIdx == Skip)
            {
                //check if there are edges beyond the skip edge in the bound and if so
                //create another LocMin and calling ProcessBound once more ...
                E = Result;
                if (LeftBoundIsForward)
                {
                    while (E.Top.Y == E.Next.Bot.Y) E = E.Next;
                    while (E != Result && E.Dx == horizontal) E = E.Prev;
                }
                else
                {
                    while (E.Top.Y == E.Prev.Bot.Y) E = E.Prev;
                    while (E != Result && E.Dx == horizontal) E = E.Next;
                }
                if (E == Result)
                {
                    if (LeftBoundIsForward) Result = E.Next;
                    else Result = E.Prev;
                }
                else
                {
                    //there are more edges in the bound beyond result starting with E
                    if (LeftBoundIsForward)
                        E = Result.Next;
                    else
                        E = Result.Prev;
                    LocalMinima locMin = new LocalMinima();
                    locMin.Next = null;
                    locMin.Y = E.Bot.Y;
                    locMin.LeftBound = null;
                    locMin.RightBound = E;
                    E.WindDelta = 0;
                    Result = ProcessBound(E, LeftBoundIsForward);
                    InsertLocalMinima(locMin);
                }
                return Result;
            }

            if (E.Dx == horizontal)
            {
                //We need to be careful with open paths because this may not be a
                //true local minima (ie E may be following a skip edge).
                //Also, consecutive horz. edges may start heading left before going right.
                if (LeftBoundIsForward) EStart = E.Prev;
                else EStart = E.Next;
                if (EStart.Dx == horizontal) //ie an adjoining horizontal skip edge
                {
                    if (EStart.Bot.X != E.Bot.X && EStart.Top.X != E.Bot.X)
                        ReverseHorizontal(E);
                }
                else if (EStart.Bot.X != E.Bot.X)
                    ReverseHorizontal(E);
            }

            EStart = E;
            if (LeftBoundIsForward)
            {
                while (Result.Top.Y == Result.Next.Bot.Y && Result.Next.OutIdx != Skip)
                    Result = Result.Next;
                if (Result.Dx == horizontal && Result.Next.OutIdx != Skip)
                {
                    //nb: at the top of a bound, horizontals are added to the bound
                    //only when the preceding edge attaches to the horizontal's left vertex
                    //unless a Skip edge is encountered when that becomes the top divide
                    Horz = Result;
                    while (Horz.Prev.Dx == horizontal) Horz = Horz.Prev;
                    if (Horz.Prev.Top.X > Result.Next.Top.X) Result = Horz.Prev;
                }
                while (E != Result)
                {
                    E.NextInLML = E.Next;
                    if (E.Dx == horizontal && E != EStart && E.Bot.X != E.Prev.Top.X)
                        ReverseHorizontal(E);
                    E = E.Next;
                }
                if (E.Dx == horizontal && E != EStart && E.Bot.X != E.Prev.Top.X)
                    ReverseHorizontal(E);
                Result = Result.Next; //move to the edge just beyond current bound
            }
            else
            {
                while (Result.Top.Y == Result.Prev.Bot.Y && Result.Prev.OutIdx != Skip)
                    Result = Result.Prev;
                if (Result.Dx == horizontal && Result.Prev.OutIdx != Skip)
                {
                    Horz = Result;
                    while (Horz.Next.Dx == horizontal) Horz = Horz.Next;
                    if (Horz.Next.Top.X == Result.Prev.Top.X ||
                        Horz.Next.Top.X > Result.Prev.Top.X) Result = Horz.Next;
                }

                while (E != Result)
                {
                    E.NextInLML = E.Prev;
                    if (E.Dx == horizontal && E != EStart && E.Bot.X != E.Next.Top.X)
                        ReverseHorizontal(E);
                    E = E.Prev;
                }
                if (E.Dx == horizontal && E != EStart && E.Bot.X != E.Next.Top.X)
                    ReverseHorizontal(E);
                Result = Result.Prev; //move to the edge just beyond current bound
            }
            return Result;
        }
        //------------------------------------------------------------------------------


        public bool AddPath(Path pg, PolyType polyType, bool Closed)
        {
            if (!Closed && polyType == PolyType.ptClip)
                throw new ClipperException("AddPath: Open paths must be subject.");

            int highI = (int)pg.Count - 1;
            if (Closed) while (highI > 0 && (pg[highI] == pg[0])) --highI;
            while (highI > 0 && (pg[highI] == pg[highI - 1])) --highI;
            if ((Closed && highI < 2) || (!Closed && highI < 1)) return false;

            //create a new edge array ...
            List<TEdge> edges = new List<TEdge>(highI + 1);
            for (int i = 0; i <= highI; i++) edges.Add(new TEdge());

            bool IsFlat = true;

            //1. Basic (first) edge initialization ...
            edges[1].Curr = pg[1];
            RangeTest(pg[0], ref m_UseFullRange);
            RangeTest(pg[highI], ref m_UseFullRange);
            InitEdge(edges[0], edges[1], edges[highI], pg[0]);
            InitEdge(edges[highI], edges[0], edges[highI - 1], pg[highI]);
            for (int i = highI - 1; i >= 1; --i)
            {
                RangeTest(pg[i], ref m_UseFullRange);
                InitEdge(edges[i], edges[i + 1], edges[i - 1], pg[i]);
            }
            TEdge eStart = edges[0];

            //2. Remove duplicate vertices, and (when closed) collinear edges ...
            TEdge E = eStart, eLoopStop = eStart;
            for (; ; )
            {
                //nb: allows matching start and end points when not Closed ...
                if (E.Curr == E.Next.Curr && (Closed || E.Next != eStart))
                {
                    if (E == E.Next) break;
                    if (E == eStart) eStart = E.Next;
                    E = RemoveEdge(E);
                    eLoopStop = E;
                    continue;
                }
                if (E.Prev == E.Next)
                    break; //only two vertices
                else if (Closed &&
                  SlopesEqual(E.Prev.Curr, E.Curr, E.Next.Curr, m_UseFullRange) &&
                  (!PreserveCollinear ||
                  !Pt2IsBetweenPt1AndPt3(E.Prev.Curr, E.Curr, E.Next.Curr)))
                {
                    //Collinear edges are allowed for open paths but in closed paths
                    //the default is to merge adjacent collinear edges into a single edge.
                    //However, if the PreserveCollinear property is enabled, only overlapping
                    //collinear edges (ie spikes) will be removed from closed paths.
                    if (E == eStart) eStart = E.Next;
                    E = RemoveEdge(E);
                    E = E.Prev;
                    eLoopStop = E;
                    continue;
                }
                E = E.Next;
                if ((E == eLoopStop) || (!Closed && E.Next == eStart)) break;
            }

            if ((!Closed && (E == E.Next)) || (Closed && (E.Prev == E.Next)))
                return false;

            if (!Closed)
            {
                m_HasOpenPaths = true;
                eStart.Prev.OutIdx = Skip;
            }

            //3. Do second stage of edge initialization ...
            E = eStart;
            do
            {
                InitEdge2(E, polyType);
                E = E.Next;
                if (IsFlat && E.Curr.Y != eStart.Curr.Y) IsFlat = false;
            }
            while (E != eStart);

            //4. Finally, add edge bounds to LocalMinima list ...

            //Totally flat paths must be handled differently when adding them
            //to LocalMinima list to avoid endless loops etc ...
            if (IsFlat)
            {
                if (Closed) return false;
                E.Prev.OutIdx = Skip;
                LocalMinima locMin = new LocalMinima();
                locMin.Next = null;
                locMin.Y = E.Bot.Y;
                locMin.LeftBound = null;
                locMin.RightBound = E;
                locMin.RightBound.Side = EdgeSide.esRight;
                locMin.RightBound.WindDelta = 0;
                for (; ; )
                {
                    if (E.Bot.X != E.Prev.Top.X) ReverseHorizontal(E);
                    if (E.Next.OutIdx == Skip) break;
                    E.NextInLML = E.Next;
                    E = E.Next;
                }
                InsertLocalMinima(locMin);
                m_edges.Add(edges);
                return true;
            }

            m_edges.Add(edges);
            bool leftBoundIsForward;
            TEdge EMin = null;

            //workaround to avoid an endless loop in the while loop below when
            //open paths have matching start and end points ...
            if (E.Prev.Bot == E.Prev.Top) E = E.Next;

            for (; ; )
            {
                E = FindNextLocMin(E);
                if (E == EMin) break;
                else if (EMin == null) EMin = E;

                //E and E.Prev now share a local minima (left aligned if horizontal).
                //Compare their slopes to find which starts which bound ...
                LocalMinima locMin = new LocalMinima();
                locMin.Next = null;
                locMin.Y = E.Bot.Y;
                if (E.Dx < E.Prev.Dx)
                {
                    locMin.LeftBound = E.Prev;
                    locMin.RightBound = E;
                    leftBoundIsForward = false; //Q.nextInLML = Q.prev
                }
                else
                {
                    locMin.LeftBound = E;
                    locMin.RightBound = E.Prev;
                    leftBoundIsForward = true; //Q.nextInLML = Q.next
                }
                locMin.LeftBound.Side = EdgeSide.esLeft;
                locMin.RightBound.Side = EdgeSide.esRight;

                if (!Closed) locMin.LeftBound.WindDelta = 0;
                else if (locMin.LeftBound.Next == locMin.RightBound)
                    locMin.LeftBound.WindDelta = -1;
                else locMin.LeftBound.WindDelta = 1;
                locMin.RightBound.WindDelta = -locMin.LeftBound.WindDelta;

                E = ProcessBound(locMin.LeftBound, leftBoundIsForward);
                if (E.OutIdx == Skip) E = ProcessBound(E, leftBoundIsForward);

                TEdge E2 = ProcessBound(locMin.RightBound, !leftBoundIsForward);
                if (E2.OutIdx == Skip) E2 = ProcessBound(E2, !leftBoundIsForward);

                if (locMin.LeftBound.OutIdx == Skip)
                    locMin.LeftBound = null;
                else if (locMin.RightBound.OutIdx == Skip)
                    locMin.RightBound = null;
                InsertLocalMinima(locMin);
                if (!leftBoundIsForward) E = E2;
            }
            return true;

        }
        //------------------------------------------------------------------------------

        public bool AddPaths(Paths ppg, PolyType polyType, bool closed)
        {
            bool result = false;
            for (int i = 0; i < ppg.Count; ++i)
                if (AddPath(ppg[i], polyType, closed)) result = true;
            return result;
        }
        //------------------------------------------------------------------------------

        internal bool Pt2IsBetweenPt1AndPt3(IntPoint pt1, IntPoint pt2, IntPoint pt3)
        {
            if ((pt1 == pt3) || (pt1 == pt2) || (pt3 == pt2)) return false;
            else if (pt1.X != pt3.X) return (pt2.X > pt1.X) == (pt2.X < pt3.X);
            else return (pt2.Y > pt1.Y) == (pt2.Y < pt3.Y);
        }
        //------------------------------------------------------------------------------

        TEdge RemoveEdge(TEdge e)
        {
            //removes e from double_linked_list (but without removing from memory)
            e.Prev.Next = e.Next;
            e.Next.Prev = e.Prev;
            TEdge result = e.Next;
            e.Prev = null; //flag as removed (see ClipperBase.Clear)
            return result;
        }
        //------------------------------------------------------------------------------

        private void SetDx(TEdge e)
        {
            e.Delta.X = (e.Top.X - e.Bot.X);
            e.Delta.Y = (e.Top.Y - e.Bot.Y);
            if (e.Delta.Y == 0) e.Dx = horizontal;
            else e.Dx = (double)(e.Delta.X) / (e.Delta.Y);
        }
        //---------------------------------------------------------------------------

        private void InsertLocalMinima(LocalMinima newLm)
        {
            if (m_MinimaList == null)
            {
                m_MinimaList = newLm;
            }
            else if (newLm.Y >= m_MinimaList.Y)
            {
                newLm.Next = m_MinimaList;
                m_MinimaList = newLm;
            }
            else
            {
                LocalMinima tmpLm = m_MinimaList;
                while (tmpLm.Next != null && (newLm.Y < tmpLm.Next.Y))
                    tmpLm = tmpLm.Next;
                newLm.Next = tmpLm.Next;
                tmpLm.Next = newLm;
            }
        }
        //------------------------------------------------------------------------------

        internal Boolean PopLocalMinima(cInt Y, out LocalMinima current)
        {
            current = m_CurrentLM;
            if (m_CurrentLM != null && m_CurrentLM.Y == Y)
            {
                m_CurrentLM = m_CurrentLM.Next;
                return true;
            }
            return false;
        }
        //------------------------------------------------------------------------------

        private void ReverseHorizontal(TEdge e)
        {
            //swap horizontal edges' top and bottom x's so they follow the natural
            //progression of the bounds - ie so their xbots will align with the
            //adjoining lower edge. [Helpful in the ProcessHorizontal() method.]
            Swap(ref e.Top.X, ref e.Bot.X);
        }
        //------------------------------------------------------------------------------

        internal virtual void Reset()
        {
            m_CurrentLM = m_MinimaList;
            if (m_CurrentLM == null) return; //ie nothing to process

            //reset all edges ...
            m_Scanbeam = null;
            LocalMinima lm = m_MinimaList;
            while (lm != null)
            {
                InsertScanbeam(lm.Y);
                TEdge e = lm.LeftBound;
                if (e != null)
                {
                    e.Curr = e.Bot;
                    e.OutIdx = Unassigned;
                }
                e = lm.RightBound;
                if (e != null)
                {
                    e.Curr = e.Bot;
                    e.OutIdx = Unassigned;
                }
                lm = lm.Next;
            }
            m_ActiveEdges = null;
        }
        //------------------------------------------------------------------------------

        public static IntRect GetBounds(Paths paths)
        {
            int i = 0, cnt = paths.Count;
            while (i < cnt && paths[i].Count == 0) i++;
            if (i == cnt) return new IntRect(0, 0, 0, 0);
            IntRect result = new IntRect();
            result.left = paths[i][0].X;
            result.right = result.left;
            result.top = paths[i][0].Y;
            result.bottom = result.top;
            for (; i < cnt; i++)
                for (int j = 0; j < paths[i].Count; j++)
                {
                    if (paths[i][j].X < result.left) result.left = paths[i][j].X;
                    else if (paths[i][j].X > result.right) result.right = paths[i][j].X;
                    if (paths[i][j].Y < result.top) result.top = paths[i][j].Y;
                    else if (paths[i][j].Y > result.bottom) result.bottom = paths[i][j].Y;
                }
            return result;
        }
        //------------------------------------------------------------------------------

        internal void InsertScanbeam(cInt Y)
        {
            //single-linked list: sorted descending, ignoring dups.
            if (m_Scanbeam == null)
            {
                m_Scanbeam = new Scanbeam();
                m_Scanbeam.Next = null;
                m_Scanbeam.Y = Y;
            }
            else if (Y > m_Scanbeam.Y)
            {
                Scanbeam newSb = new Scanbeam();
                newSb.Y = Y;
                newSb.Next = m_Scanbeam;
                m_Scanbeam = newSb;
            }
            else
            {
                Scanbeam sb2 = m_Scanbeam;
                while (sb2.Next != null && (Y <= sb2.Next.Y)) sb2 = sb2.Next;
                if (Y == sb2.Y) return; //ie ignores duplicates
                Scanbeam newSb = new Scanbeam();
                newSb.Y = Y;
                newSb.Next = sb2.Next;
                sb2.Next = newSb;
            }
        }
        //------------------------------------------------------------------------------

        internal Boolean PopScanbeam(out cInt Y)
        {
            if (m_Scanbeam == null)
            {
                Y = 0;
                return false;
            }
            Y = m_Scanbeam.Y;
            m_Scanbeam = m_Scanbeam.Next;
            return true;
        }
        //------------------------------------------------------------------------------

        internal Boolean LocalMinimaPending()
        {
            return (m_CurrentLM != null);
        }
        //------------------------------------------------------------------------------

        internal OutRec CreateOutRec()
        {
            OutRec result = new OutRec();
            result.Idx = Unassigned;
            result.IsHole = false;
            result.IsOpen = false;
            result.FirstLeft = null;
            result.Pts = null;
            result.BottomPt = null;
            result.PolyNode = null;
            m_PolyOuts.Add(result);
            result.Idx = m_PolyOuts.Count - 1;
            return result;
        }
        //------------------------------------------------------------------------------

        internal void DisposeOutRec(int index)
        {
            OutRec outRec = m_PolyOuts[index];
            outRec.Pts = null;
            outRec = null;
            m_PolyOuts[index] = null;
        }
        //------------------------------------------------------------------------------

        internal void UpdateEdgeIntoAEL(ref TEdge e)
        {
            if (e.NextInLML == null)
                throw new ClipperException("UpdateEdgeIntoAEL: invalid call");
            TEdge AelPrev = e.PrevInAEL;
            TEdge AelNext = e.NextInAEL;
            e.NextInLML.OutIdx = e.OutIdx;
            if (AelPrev != null)
                AelPrev.NextInAEL = e.NextInLML;
            else m_ActiveEdges = e.NextInLML;
            if (AelNext != null)
                AelNext.PrevInAEL = e.NextInLML;
            e.NextInLML.Side = e.Side;
            e.NextInLML.WindDelta = e.WindDelta;
            e.NextInLML.WindCnt = e.WindCnt;
            e.NextInLML.WindCnt2 = e.WindCnt2;
            e = e.NextInLML;
            e.Curr = e.Bot;
            e.PrevInAEL = AelPrev;
            e.NextInAEL = AelNext;
            if (!IsHorizontal(e)) InsertScanbeam(e.Top.Y);
        }
        //------------------------------------------------------------------------------

        internal void SwapPositionsInAEL(TEdge edge1, TEdge edge2)
        {
            //check that one or other edge hasn't already been removed from AEL ...
            if (edge1.NextInAEL == edge1.PrevInAEL ||
              edge2.NextInAEL == edge2.PrevInAEL) return;

            if (edge1.NextInAEL == edge2)
            {
                TEdge next = edge2.NextInAEL;
                if (next != null)
                    next.PrevInAEL = edge1;
                TEdge prev = edge1.PrevInAEL;
                if (prev != null)
                    prev.NextInAEL = edge2;
                edge2.PrevInAEL = prev;
                edge2.NextInAEL = edge1;
                edge1.PrevInAEL = edge2;
                edge1.NextInAEL = next;
            }
            else if (edge2.NextInAEL == edge1)
            {
                TEdge next = edge1.NextInAEL;
                if (next != null)
                    next.PrevInAEL = edge2;
                TEdge prev = edge2.PrevInAEL;
                if (prev != null)
                    prev.NextInAEL = edge1;
                edge1.PrevInAEL = prev;
                edge1.NextInAEL = edge2;
                edge2.PrevInAEL = edge1;
                edge2.NextInAEL = next;
            }
            else
            {
                TEdge next = edge1.NextInAEL;
                TEdge prev = edge1.PrevInAEL;
                edge1.NextInAEL = edge2.NextInAEL;
                if (edge1.NextInAEL != null)
                    edge1.NextInAEL.PrevInAEL = edge1;
                edge1.PrevInAEL = edge2.PrevInAEL;
                if (edge1.PrevInAEL != null)
                    edge1.PrevInAEL.NextInAEL = edge1;
                edge2.NextInAEL = next;
                if (edge2.NextInAEL != null)
                    edge2.NextInAEL.PrevInAEL = edge2;
                edge2.PrevInAEL = prev;
                if (edge2.PrevInAEL != null)
                    edge2.PrevInAEL.NextInAEL = edge2;
            }

            if (edge1.PrevInAEL == null)
                m_ActiveEdges = edge1;
            else if (edge2.PrevInAEL == null)
                m_ActiveEdges = edge2;
        }
        //------------------------------------------------------------------------------

        internal void DeleteFromAEL(TEdge e)
        {
            TEdge AelPrev = e.PrevInAEL;
            TEdge AelNext = e.NextInAEL;
            if (AelPrev == null && AelNext == null && (e != m_ActiveEdges))
                return; //already deleted
            if (AelPrev != null)
                AelPrev.NextInAEL = AelNext;
            else m_ActiveEdges = AelNext;
            if (AelNext != null)
                AelNext.PrevInAEL = AelPrev;
            e.NextInAEL = null;
            e.PrevInAEL = null;
        }
        //------------------------------------------------------------------------------

    } //end ClipperBase

    public class Clipper : ClipperBase
    {
        //InitOptions that can be passed to the constructor ...
        public const int ioReverseSolution = 1;
        public const int ioStrictlySimple = 2;
        public const int ioPreserveCollinear = 4;

        private ClipType m_ClipType;
        private Maxima m_Maxima;
        private TEdge m_SortedEdges;
        private List<IntersectNode> m_IntersectList;
        IComparer<IntersectNode> m_IntersectNodeComparer;
        private bool m_ExecuteLocked;
        private PolyFillType m_ClipFillType;
        private PolyFillType m_SubjFillType;
        private List<Join> m_Joins;
        private List<Join> m_GhostJoins;
        private bool m_UsingPolyTree;
        public Clipper(int InitOptions = 0) : base() //constructor
        {
            m_Scanbeam = null;
            m_Maxima = null;
            m_ActiveEdges = null;
            m_SortedEdges = null;
            m_IntersectList = new List<IntersectNode>();
            m_IntersectNodeComparer = new MyIntersectNodeSort();
            m_ExecuteLocked = false;
            m_UsingPolyTree = false;
            m_PolyOuts = new List<OutRec>();
            m_Joins = new List<Join>();
            m_GhostJoins = new List<Join>();
            ReverseSolution = (ioReverseSolution & InitOptions) != 0;
            StrictlySimple = (ioStrictlySimple & InitOptions) != 0;
            PreserveCollinear = (ioPreserveCollinear & InitOptions) != 0;
        }
        //------------------------------------------------------------------------------

        private void InsertMaxima(cInt X)
        {
            //double-linked list: sorted ascending, ignoring dups.
            Maxima newMax = new Maxima();
            newMax.X = X;
            if (m_Maxima == null)
            {
                m_Maxima = newMax;
                m_Maxima.Next = null;
                m_Maxima.Prev = null;
            }
            else if (X < m_Maxima.X)
            {
                newMax.Next = m_Maxima;
                newMax.Prev = null;
                m_Maxima = newMax;
            }
            else
            {
                Maxima m = m_Maxima;
                while (m.Next != null && (X >= m.Next.X)) m = m.Next;
                if (X == m.X) return; //ie ignores duplicates (& CG to clean up newMax)
                                      //insert newMax between m and m.Next ...
                newMax.Next = m.Next;
                newMax.Prev = m;
                if (m.Next != null) m.Next.Prev = newMax;
                m.Next = newMax;
            }
        }
        //------------------------------------------------------------------------------

        public bool ReverseSolution
        {
            get;
            set;
        }
        //------------------------------------------------------------------------------

        public bool StrictlySimple
        {
            get;
            set;
        }
        //------------------------------------------------------------------------------

        public bool Execute(ClipType clipType, Paths solution,
            PolyFillType FillType = PolyFillType.pftEvenOdd)
        {
            return Execute(clipType, solution, FillType, FillType);
        }
        //------------------------------------------------------------------------------

        public bool Execute(ClipType clipType, PolyTree polytree,
            PolyFillType FillType = PolyFillType.pftEvenOdd)
        {
            return Execute(clipType, polytree, FillType, FillType);
        }
        //------------------------------------------------------------------------------

        public bool Execute(ClipType clipType, Paths solution,
            PolyFillType subjFillType, PolyFillType clipFillType)
        {
            if (m_ExecuteLocked) return false;
            if (m_HasOpenPaths) throw
              new ClipperException("Error: PolyTree struct is needed for open path clipping.");

            m_ExecuteLocked = true;
            solution.Clear();
            m_SubjFillType = subjFillType;
            m_ClipFillType = clipFillType;
            m_ClipType = clipType;
            m_UsingPolyTree = false;
            bool succeeded;
            try
            {
                succeeded = ExecuteInternal();
                //build the return polygons ...
                if (succeeded) BuildResult(solution);
            }
            finally
            {
                DisposeAllPolyPts();
                m_ExecuteLocked = false;
            }
            return succeeded;
        }
        //------------------------------------------------------------------------------

        public bool Execute(ClipType clipType, PolyTree polytree,
            PolyFillType subjFillType, PolyFillType clipFillType)
        {
            if (m_ExecuteLocked) return false;
            m_ExecuteLocked = true;
            m_SubjFillType = subjFillType;
            m_ClipFillType = clipFillType;
            m_ClipType = clipType;
            m_UsingPolyTree = true;
            bool succeeded;
            try
            {
                succeeded = ExecuteInternal();
                //build the return polygons ...
                if (succeeded) BuildResult2(polytree);
            }
            finally
            {
                DisposeAllPolyPts();
                m_ExecuteLocked = false;
            }
            return succeeded;
        }
        //------------------------------------------------------------------------------

        internal void FixHoleLinkage(OutRec outRec)
        {
            //skip if an outermost polygon or
            //already already points to the correct FirstLeft ...
            if (outRec.FirstLeft == null ||
                  (outRec.IsHole != outRec.FirstLeft.IsHole &&
                  outRec.FirstLeft.Pts != null)) return;

            OutRec orfl = outRec.FirstLeft;
            while (orfl != null && ((orfl.IsHole == outRec.IsHole) || orfl.Pts == null))
                orfl = orfl.FirstLeft;
            outRec.FirstLeft = orfl;
        }
        //------------------------------------------------------------------------------

        private bool ExecuteInternal()
        {
            try
            {
                Reset();
                m_SortedEdges = null;
                m_Maxima = null;

                cInt botY, topY;
                if (!PopScanbeam(out botY)) return false;
                InsertLocalMinimaIntoAEL(botY);
                while (PopScanbeam(out topY) || LocalMinimaPending())
                {
                    ProcessHorizontals();
                    m_GhostJoins.Clear();
                    if (!ProcessIntersections(topY)) return false;
                    ProcessEdgesAtTopOfScanbeam(topY);
                    botY = topY;
                    InsertLocalMinimaIntoAEL(botY);
                }

                //fix orientations ...
                foreach (OutRec outRec in m_PolyOuts)
                {
                    if (outRec.Pts == null || outRec.IsOpen) continue;
                    if ((outRec.IsHole ^ ReverseSolution) == (Area(outRec) > 0))
                        ReversePolyPtLinks(outRec.Pts);
                }

                JoinCommonEdges();

                foreach (OutRec outRec in m_PolyOuts)
                {
                    if (outRec.Pts == null)
                        continue;
                    else if (outRec.IsOpen)
                        FixupOutPolyline(outRec);
                    else
                        FixupOutPolygon(outRec);
                }

                if (StrictlySimple) DoSimplePolygons();
                return true;
            }
            //catch { return false; }
            finally
            {
                m_Joins.Clear();
                m_GhostJoins.Clear();
            }
        }
        //------------------------------------------------------------------------------

        private void DisposeAllPolyPts()
        {
            for (int i = 0; i < m_PolyOuts.Count; ++i) DisposeOutRec(i);
            m_PolyOuts.Clear();
        }
        //------------------------------------------------------------------------------

        private void AddJoin(OutPt Op1, OutPt Op2, IntPoint OffPt)
        {
            Join j = new Join();
            j.OutPt1 = Op1;
            j.OutPt2 = Op2;
            j.OffPt = OffPt;
            m_Joins.Add(j);
        }
        //------------------------------------------------------------------------------

        private void AddGhostJoin(OutPt Op, IntPoint OffPt)
        {
            Join j = new Join();
            j.OutPt1 = Op;
            j.OffPt = OffPt;
            m_GhostJoins.Add(j);
        }
        //------------------------------------------------------------------------------


        private void InsertLocalMinimaIntoAEL(cInt botY)
        {
            LocalMinima lm;
            while (PopLocalMinima(botY, out lm))
            {
                TEdge lb = lm.LeftBound;
                TEdge rb = lm.RightBound;

                OutPt Op1 = null;
                if (lb == null)
                {
                    InsertEdgeIntoAEL(rb, null);
                    SetWindingCount(rb);
                    if (IsContributing(rb))
                        Op1 = AddOutPt(rb, rb.Bot);
                }
                else if (rb == null)
                {
                    InsertEdgeIntoAEL(lb, null);
                    SetWindingCount(lb);
                    if (IsContributing(lb))
                        Op1 = AddOutPt(lb, lb.Bot);
                    InsertScanbeam(lb.Top.Y);
                }
                else
                {
                    InsertEdgeIntoAEL(lb, null);
                    InsertEdgeIntoAEL(rb, lb);
                    SetWindingCount(lb);
                    rb.WindCnt = lb.WindCnt;
                    rb.WindCnt2 = lb.WindCnt2;
                    if (IsContributing(lb))
                        Op1 = AddLocalMinPoly(lb, rb, lb.Bot);
                    InsertScanbeam(lb.Top.Y);
                }

                if (rb != null)
                {
                    if (IsHorizontal(rb))
                    {
                        if (rb.NextInLML != null)
                            InsertScanbeam(rb.NextInLML.Top.Y);
                        AddEdgeToSEL(rb);
                    }
                    else
                        InsertScanbeam(rb.Top.Y);
                }

                if (lb == null || rb == null) continue;

                //if output polygons share an Edge with a horizontal rb, they'll need joining later ...
                if (Op1 != null && IsHorizontal(rb) &&
                  m_GhostJoins.Count > 0 && rb.WindDelta != 0)
                {
                    for (int i = 0; i < m_GhostJoins.Count; i++)
                    {
                        //if the horizontal Rb and a 'ghost' horizontal overlap, then convert
                        //the 'ghost' join to a real join ready for later ...
                        Join j = m_GhostJoins[i];
                        if (HorzSegmentsOverlap(j.OutPt1.Pt.X, j.OffPt.X, rb.Bot.X, rb.Top.X))
                            AddJoin(j.OutPt1, Op1, j.OffPt);
                    }
                }

                if (lb.OutIdx >= 0 && lb.PrevInAEL != null &&
                  lb.PrevInAEL.Curr.X == lb.Bot.X &&
                  lb.PrevInAEL.OutIdx >= 0 &&
                  SlopesEqual(lb.PrevInAEL.Curr, lb.PrevInAEL.Top, lb.Curr, lb.Top, m_UseFullRange) &&
                  lb.WindDelta != 0 && lb.PrevInAEL.WindDelta != 0)
                {
                    OutPt Op2 = AddOutPt(lb.PrevInAEL, lb.Bot);
                    AddJoin(Op1, Op2, lb.Top);
                }

                if (lb.NextInAEL != rb)
                {

                    if (rb.OutIdx >= 0 && rb.PrevInAEL.OutIdx >= 0 &&
                      SlopesEqual(rb.PrevInAEL.Curr, rb.PrevInAEL.Top, rb.Curr, rb.Top, m_UseFullRange) &&
                      rb.WindDelta != 0 && rb.PrevInAEL.WindDelta != 0)
                    {
                        OutPt Op2 = AddOutPt(rb.PrevInAEL, rb.Bot);
                        AddJoin(Op1, Op2, rb.Top);
                    }

                    TEdge e = lb.NextInAEL;
                    if (e != null)
                        while (e != rb)
                        {
                            //nb: For calculating winding counts etc, IntersectEdges() assumes
                            //that param1 will be to the right of param2 ABOVE the intersection ...
                            IntersectEdges(rb, e, lb.Curr); //order important here
                            e = e.NextInAEL;
                        }
                }
            }
        }
        //------------------------------------------------------------------------------

        private void InsertEdgeIntoAEL(TEdge edge, TEdge startEdge)
        {
            if (m_ActiveEdges == null)
            {
                edge.PrevInAEL = null;
                edge.NextInAEL = null;
                m_ActiveEdges = edge;
            }
            else if (startEdge == null && E2InsertsBeforeE1(m_ActiveEdges, edge))
            {
                edge.PrevInAEL = null;
                edge.NextInAEL = m_ActiveEdges;
                m_ActiveEdges.PrevInAEL = edge;
                m_ActiveEdges = edge;
            }
            else
            {
                if (startEdge == null) startEdge = m_ActiveEdges;
                while (startEdge.NextInAEL != null &&
                  !E2InsertsBeforeE1(startEdge.NextInAEL, edge))
                    startEdge = startEdge.NextInAEL;
                edge.NextInAEL = startEdge.NextInAEL;
                if (startEdge.NextInAEL != null) startEdge.NextInAEL.PrevInAEL = edge;
                edge.PrevInAEL = startEdge;
                startEdge.NextInAEL = edge;
            }
        }
        //----------------------------------------------------------------------

        private bool E2InsertsBeforeE1(TEdge e1, TEdge e2)
        {
            if (e2.Curr.X == e1.Curr.X)
            {
                if (e2.Top.Y > e1.Top.Y)
                    return e2.Top.X < TopX(e1, e2.Top.Y);
                else return e1.Top.X > TopX(e2, e1.Top.Y);
            }
            else return e2.Curr.X < e1.Curr.X;
        }
        //------------------------------------------------------------------------------

        private bool IsEvenOddFillType(TEdge edge)
        {
            if (edge.PolyTyp == PolyType.ptSubject)
                return m_SubjFillType == PolyFillType.pftEvenOdd;
            else
                return m_ClipFillType == PolyFillType.pftEvenOdd;
        }
        //------------------------------------------------------------------------------

        private bool IsEvenOddAltFillType(TEdge edge)
        {
            if (edge.PolyTyp == PolyType.ptSubject)
                return m_ClipFillType == PolyFillType.pftEvenOdd;
            else
                return m_SubjFillType == PolyFillType.pftEvenOdd;
        }
        //------------------------------------------------------------------------------

        private bool IsContributing(TEdge edge)
        {
            PolyFillType pft, pft2;
            if (edge.PolyTyp == PolyType.ptSubject)
            {
                pft = m_SubjFillType;
                pft2 = m_ClipFillType;
            }
            else
            {
                pft = m_ClipFillType;
                pft2 = m_SubjFillType;
            }

            switch (pft)
            {
                case PolyFillType.pftEvenOdd:
                    //return false if a subj line has been flagged as inside a subj polygon
                    if (edge.WindDelta == 0 && edge.WindCnt != 1) return false;
                    break;
                case PolyFillType.pftNonZero:
                    if (Math.Abs(edge.WindCnt) != 1) return false;
                    break;
                case PolyFillType.pftPositive:
                    if (edge.WindCnt != 1) return false;
                    break;
                default: //PolyFillType.pftNegative
                    if (edge.WindCnt != -1) return false;
                    break;
            }

            switch (m_ClipType)
            {
                case ClipType.ctIntersection:
                    switch (pft2)
                    {
                        case PolyFillType.pftEvenOdd:
                        case PolyFillType.pftNonZero:
                            return (edge.WindCnt2 != 0);
                        case PolyFillType.pftPositive:
                            return (edge.WindCnt2 > 0);
                        default:
                            return (edge.WindCnt2 < 0);
                    }
                case ClipType.ctUnion:
                    switch (pft2)
                    {
                        case PolyFillType.pftEvenOdd:
                        case PolyFillType.pftNonZero:
                            return (edge.WindCnt2 == 0);
                        case PolyFillType.pftPositive:
                            return (edge.WindCnt2 <= 0);
                        default:
                            return (edge.WindCnt2 >= 0);
                    }
                case ClipType.ctDifference:
                    if (edge.PolyTyp == PolyType.ptSubject)
                        switch (pft2)
                        {
                            case PolyFillType.pftEvenOdd:
                            case PolyFillType.pftNonZero:
                                return (edge.WindCnt2 == 0);
                            case PolyFillType.pftPositive:
                                return (edge.WindCnt2 <= 0);
                            default:
                                return (edge.WindCnt2 >= 0);
                        }
                    else
                        switch (pft2)
                        {
                            case PolyFillType.pftEvenOdd:
                            case PolyFillType.pftNonZero:
                                return (edge.WindCnt2 != 0);
                            case PolyFillType.pftPositive:
                                return (edge.WindCnt2 > 0);
                            default:
                                return (edge.WindCnt2 < 0);
                        }
                case ClipType.ctXor:
                    if (edge.WindDelta == 0) //XOr always contributing unless open
                        switch (pft2)
                        {
                            case PolyFillType.pftEvenOdd:
                            case PolyFillType.pftNonZero:
                                return (edge.WindCnt2 == 0);
                            case PolyFillType.pftPositive:
                                return (edge.WindCnt2 <= 0);
                            default:
                                return (edge.WindCnt2 >= 0);
                        }
                    else
                        return true;
            }
            return true;
        }
        //------------------------------------------------------------------------------

        private void SetWindingCount(TEdge edge)
        {
            TEdge e = edge.PrevInAEL;
            //find the edge of the same polytype that immediately preceeds 'edge' in AEL
            while (e != null && ((e.PolyTyp != edge.PolyTyp) || (e.WindDelta == 0))) e = e.PrevInAEL;
            if (e == null)
            {
                PolyFillType pft;
                pft = (edge.PolyTyp == PolyType.ptSubject ? m_SubjFillType : m_ClipFillType);
                if (edge.WindDelta == 0) edge.WindCnt = (pft == PolyFillType.pftNegative ? -1 : 1);
                else edge.WindCnt = edge.WindDelta;
                edge.WindCnt2 = 0;
                e = m_ActiveEdges; //ie get ready to calc WindCnt2
            }
            else if (edge.WindDelta == 0 && m_ClipType != ClipType.ctUnion)
            {
                edge.WindCnt = 1;
                edge.WindCnt2 = e.WindCnt2;
                e = e.NextInAEL; //ie get ready to calc WindCnt2
            }
            else if (IsEvenOddFillType(edge))
            {
                //EvenOdd filling ...
                if (edge.WindDelta == 0)
                {
                    //are we inside a subj polygon ...
                    bool Inside = true;
                    TEdge e2 = e.PrevInAEL;
                    while (e2 != null)
                    {
                        if (e2.PolyTyp == e.PolyTyp && e2.WindDelta != 0)
                            Inside = !Inside;
                        e2 = e2.PrevInAEL;
                    }
                    edge.WindCnt = (Inside ? 0 : 1);
                }
                else
                {
                    edge.WindCnt = edge.WindDelta;
                }
                edge.WindCnt2 = e.WindCnt2;
                e = e.NextInAEL; //ie get ready to calc WindCnt2
            }
            else
            {
                //nonZero, Positive or Negative filling ...
                if (e.WindCnt * e.WindDelta < 0)
                {
                    //prev edge is 'decreasing' WindCount (WC) toward zero
                    //so we're outside the previous polygon ...
                    if (Math.Abs(e.WindCnt) > 1)
                    {
                        //outside prev poly but still inside another.
                        //when reversing direction of prev poly use the same WC 
                        if (e.WindDelta * edge.WindDelta < 0) edge.WindCnt = e.WindCnt;
                        //otherwise continue to 'decrease' WC ...
                        else edge.WindCnt = e.WindCnt + edge.WindDelta;
                    }
                    else
                        //now outside all polys of same polytype so set own WC ...
                        edge.WindCnt = (edge.WindDelta == 0 ? 1 : edge.WindDelta);
                }
                else
                {
                    //prev edge is 'increasing' WindCount (WC) away from zero
                    //so we're inside the previous polygon ...
                    if (edge.WindDelta == 0)
                        edge.WindCnt = (e.WindCnt < 0 ? e.WindCnt - 1 : e.WindCnt + 1);
                    //if wind direction is reversing prev then use same WC
                    else if (e.WindDelta * edge.WindDelta < 0)
                        edge.WindCnt = e.WindCnt;
                    //otherwise add to WC ...
                    else edge.WindCnt = e.WindCnt + edge.WindDelta;
                }
                edge.WindCnt2 = e.WindCnt2;
                e = e.NextInAEL; //ie get ready to calc WindCnt2
            }

            //update WindCnt2 ...
            if (IsEvenOddAltFillType(edge))
            {
                //EvenOdd filling ...
                while (e != edge)
                {
                    if (e.WindDelta != 0)
                        edge.WindCnt2 = (edge.WindCnt2 == 0 ? 1 : 0);
                    e = e.NextInAEL;
                }
            }
            else
            {
                //nonZero, Positive or Negative filling ...
                while (e != edge)
                {
                    edge.WindCnt2 += e.WindDelta;
                    e = e.NextInAEL;
                }
            }
        }
        //------------------------------------------------------------------------------

        private void AddEdgeToSEL(TEdge edge)
        {
            //SEL pointers in PEdge are use to build transient lists of horizontal edges.
            //However, since we don't need to worry about processing order, all additions
            //are made to the front of the list ...
            if (m_SortedEdges == null)
            {
                m_SortedEdges = edge;
                edge.PrevInSEL = null;
                edge.NextInSEL = null;
            }
            else
            {
                edge.NextInSEL = m_SortedEdges;
                edge.PrevInSEL = null;
                m_SortedEdges.PrevInSEL = edge;
                m_SortedEdges = edge;
            }
        }
        //------------------------------------------------------------------------------

        internal Boolean PopEdgeFromSEL(out TEdge e)
        {
            //Pop edge from front of SEL (ie SEL is a FILO list)
            e = m_SortedEdges;
            if (e == null) return false;
            TEdge oldE = e;
            m_SortedEdges = e.NextInSEL;
            if (m_SortedEdges != null) m_SortedEdges.PrevInSEL = null;
            oldE.NextInSEL = null;
            oldE.PrevInSEL = null;
            return true;
        }
        //------------------------------------------------------------------------------

        private void CopyAELToSEL()
        {
            TEdge e = m_ActiveEdges;
            m_SortedEdges = e;
            while (e != null)
            {
                e.PrevInSEL = e.PrevInAEL;
                e.NextInSEL = e.NextInAEL;
                e = e.NextInAEL;
            }
        }
        //------------------------------------------------------------------------------

        private void SwapPositionsInSEL(TEdge edge1, TEdge edge2)
        {
            if (edge1.NextInSEL == null && edge1.PrevInSEL == null)
                return;
            if (edge2.NextInSEL == null && edge2.PrevInSEL == null)
                return;

            if (edge1.NextInSEL == edge2)
            {
                TEdge next = edge2.NextInSEL;
                if (next != null)
                    next.PrevInSEL = edge1;
                TEdge prev = edge1.PrevInSEL;
                if (prev != null)
                    prev.NextInSEL = edge2;
                edge2.PrevInSEL = prev;
                edge2.NextInSEL = edge1;
                edge1.PrevInSEL = edge2;
                edge1.NextInSEL = next;
            }
            else if (edge2.NextInSEL == edge1)
            {
                TEdge next = edge1.NextInSEL;
                if (next != null)
                    next.PrevInSEL = edge2;
                TEdge prev = edge2.PrevInSEL;
                if (prev != null)
                    prev.NextInSEL = edge1;
                edge1.PrevInSEL = prev;
                edge1.NextInSEL = edge2;
                edge2.PrevInSEL = edge1;
                edge2.NextInSEL = next;
            }
            else
            {
                TEdge next = edge1.NextInSEL;
                TEdge prev = edge1.PrevInSEL;
                edge1.NextInSEL = edge2.NextInSEL;
                if (edge1.NextInSEL != null)
                    edge1.NextInSEL.PrevInSEL = edge1;
                edge1.PrevInSEL = edge2.PrevInSEL;
                if (edge1.PrevInSEL != null)
                    edge1.PrevInSEL.NextInSEL = edge1;
                edge2.NextInSEL = next;
                if (edge2.NextInSEL != null)
                    edge2.NextInSEL.PrevInSEL = edge2;
                edge2.PrevInSEL = prev;
                if (edge2.PrevInSEL != null)
                    edge2.PrevInSEL.NextInSEL = edge2;
            }

            if (edge1.PrevInSEL == null)
                m_SortedEdges = edge1;
            else if (edge2.PrevInSEL == null)
                m_SortedEdges = edge2;
        }
        //------------------------------------------------------------------------------


        private void AddLocalMaxPoly(TEdge e1, TEdge e2, IntPoint pt)
        {
            AddOutPt(e1, pt);
            if (e2.WindDelta == 0) AddOutPt(e2, pt);
            if (e1.OutIdx == e2.OutIdx)
            {
                e1.OutIdx = Unassigned;
                e2.OutIdx = Unassigned;
            }
            else if (e1.OutIdx < e2.OutIdx)
                AppendPolygon(e1, e2);
            else
                AppendPolygon(e2, e1);
        }
        //------------------------------------------------------------------------------

        private OutPt AddLocalMinPoly(TEdge e1, TEdge e2, IntPoint pt)
        {
            OutPt result;
            TEdge e, prevE;
            if (IsHorizontal(e2) || (e1.Dx > e2.Dx))
            {
                result = AddOutPt(e1, pt);
                e2.OutIdx = e1.OutIdx;
                e1.Side = EdgeSide.esLeft;
                e2.Side = EdgeSide.esRight;
                e = e1;
                if (e.PrevInAEL == e2)
                    prevE = e2.PrevInAEL;
                else
                    prevE = e.PrevInAEL;
            }
            else
            {
                result = AddOutPt(e2, pt);
                e1.OutIdx = e2.OutIdx;
                e1.Side = EdgeSide.esRight;
                e2.Side = EdgeSide.esLeft;
                e = e2;
                if (e.PrevInAEL == e1)
                    prevE = e1.PrevInAEL;
                else
                    prevE = e.PrevInAEL;
            }

            if (prevE != null && prevE.OutIdx >= 0 && prevE.Top.Y < pt.Y && e.Top.Y < pt.Y)
            {
                cInt xPrev = TopX(prevE, pt.Y);
                cInt xE = TopX(e, pt.Y);
                if ((xPrev == xE) && (e.WindDelta != 0) && (prevE.WindDelta != 0) &&
                  SlopesEqual(new IntPoint(xPrev, pt.Y), prevE.Top, new IntPoint(xE, pt.Y), e.Top, m_UseFullRange))
                {
                    OutPt outPt = AddOutPt(prevE, pt);
                    AddJoin(result, outPt, e.Top);
                }
            }
            return result;
        }
        //------------------------------------------------------------------------------

        private OutPt AddOutPt(TEdge e, IntPoint pt)
        {
            if (e.OutIdx < 0)
            {
                OutRec outRec = CreateOutRec();
                outRec.IsOpen = (e.WindDelta == 0);
                OutPt newOp = new OutPt();
                outRec.Pts = newOp;
                newOp.Idx = outRec.Idx;
                newOp.Pt = pt;
                newOp.Next = newOp;
                newOp.Prev = newOp;
                if (!outRec.IsOpen)
                    SetHoleState(e, outRec);
                e.OutIdx = outRec.Idx; //nb: do this after SetZ !
                return newOp;
            }
            else
            {
                OutRec outRec = m_PolyOuts[e.OutIdx];
                //OutRec.Pts is the 'Left-most' point & OutRec.Pts.Prev is the 'Right-most'
                OutPt op = outRec.Pts;
                bool ToFront = (e.Side == EdgeSide.esLeft);
                if (ToFront && pt == op.Pt) return op;
                else if (!ToFront && pt == op.Prev.Pt) return op.Prev;

                OutPt newOp = new OutPt();
                newOp.Idx = outRec.Idx;
                newOp.Pt = pt;
                newOp.Next = op;
                newOp.Prev = op.Prev;
                newOp.Prev.Next = newOp;
                op.Prev = newOp;
                if (ToFront) outRec.Pts = newOp;
                return newOp;
            }
        }
        //------------------------------------------------------------------------------

        private OutPt GetLastOutPt(TEdge e)
        {
            OutRec outRec = m_PolyOuts[e.OutIdx];
            if (e.Side == EdgeSide.esLeft)
                return outRec.Pts;
            else
                return outRec.Pts.Prev;
        }
        //------------------------------------------------------------------------------

        internal void SwapPoints(ref IntPoint pt1, ref IntPoint pt2)
        {
            IntPoint tmp = new IntPoint(pt1);
            pt1 = pt2;
            pt2 = tmp;
        }
        //------------------------------------------------------------------------------

        private bool HorzSegmentsOverlap(cInt seg1a, cInt seg1b, cInt seg2a, cInt seg2b)
        {
            if (seg1a > seg1b) Swap(ref seg1a, ref seg1b);
            if (seg2a > seg2b) Swap(ref seg2a, ref seg2b);
            return (seg1a < seg2b) && (seg2a < seg1b);
        }
        //------------------------------------------------------------------------------

        private void SetHoleState(TEdge e, OutRec outRec)
        {
            TEdge e2 = e.PrevInAEL;
            TEdge eTmp = null;
            while (e2 != null)
            {
                if (e2.OutIdx >= 0 && e2.WindDelta != 0)
                {
                    if (eTmp == null)
                        eTmp = e2;
                    else if (eTmp.OutIdx == e2.OutIdx)
                        eTmp = null; //paired               
                }
                e2 = e2.PrevInAEL;
            }

            if (eTmp == null)
            {
                outRec.FirstLeft = null;
                outRec.IsHole = false;
            }
            else
            {
                outRec.FirstLeft = m_PolyOuts[eTmp.OutIdx];
                outRec.IsHole = !outRec.FirstLeft.IsHole;
            }
        }
        //------------------------------------------------------------------------------

        private double GetDx(IntPoint pt1, IntPoint pt2)
        {
            if (pt1.Y == pt2.Y) return horizontal;
            else return (double)(pt2.X - pt1.X) / (pt2.Y - pt1.Y);
        }
        //---------------------------------------------------------------------------

        private bool FirstIsBottomPt(OutPt btmPt1, OutPt btmPt2)
        {
            OutPt p = btmPt1.Prev;
            while ((p.Pt == btmPt1.Pt) && (p != btmPt1)) p = p.Prev;
            double dx1p = Math.Abs(GetDx(btmPt1.Pt, p.Pt));
            p = btmPt1.Next;
            while ((p.Pt == btmPt1.Pt) && (p != btmPt1)) p = p.Next;
            double dx1n = Math.Abs(GetDx(btmPt1.Pt, p.Pt));

            p = btmPt2.Prev;
            while ((p.Pt == btmPt2.Pt) && (p != btmPt2)) p = p.Prev;
            double dx2p = Math.Abs(GetDx(btmPt2.Pt, p.Pt));
            p = btmPt2.Next;
            while ((p.Pt == btmPt2.Pt) && (p != btmPt2)) p = p.Next;
            double dx2n = Math.Abs(GetDx(btmPt2.Pt, p.Pt));

            if (Math.Max(dx1p, dx1n) == Math.Max(dx2p, dx2n) &&
              Math.Min(dx1p, dx1n) == Math.Min(dx2p, dx2n))
                return Area(btmPt1) > 0; //if otherwise identical use orientation
            else
                return (dx1p >= dx2p && dx1p >= dx2n) || (dx1n >= dx2p && dx1n >= dx2n);
        }
        //------------------------------------------------------------------------------

        private OutPt GetBottomPt(OutPt pp)
        {
            OutPt dups = null;
            OutPt p = pp.Next;
            while (p != pp)
            {
                if (p.Pt.Y > pp.Pt.Y)
                {
                    pp = p;
                    dups = null;
                }
                else if (p.Pt.Y == pp.Pt.Y && p.Pt.X <= pp.Pt.X)
                {
                    if (p.Pt.X < pp.Pt.X)
                    {
                        dups = null;
                        pp = p;
                    }
                    else
                    {
                        if (p.Next != pp && p.Prev != pp) dups = p;
                    }
                }
                p = p.Next;
            }
            if (dups != null)
            {
                //there appears to be at least 2 vertices at bottomPt so ...
                while (dups != p)
                {
                    if (!FirstIsBottomPt(p, dups)) pp = dups;
                    dups = dups.Next;
                    while (dups.Pt != pp.Pt) dups = dups.Next;
                }
            }
            return pp;
        }
        //------------------------------------------------------------------------------

        private OutRec GetLowermostRec(OutRec outRec1, OutRec outRec2)
        {
            //work out which polygon fragment has the correct hole state ...
            if (outRec1.BottomPt == null)
                outRec1.BottomPt = GetBottomPt(outRec1.Pts);
            if (outRec2.BottomPt == null)
                outRec2.BottomPt = GetBottomPt(outRec2.Pts);
            OutPt bPt1 = outRec1.BottomPt;
            OutPt bPt2 = outRec2.BottomPt;
            if (bPt1.Pt.Y > bPt2.Pt.Y) return outRec1;
            else if (bPt1.Pt.Y < bPt2.Pt.Y) return outRec2;
            else if (bPt1.Pt.X < bPt2.Pt.X) return outRec1;
            else if (bPt1.Pt.X > bPt2.Pt.X) return outRec2;
            else if (bPt1.Next == bPt1) return outRec2;
            else if (bPt2.Next == bPt2) return outRec1;
            else if (FirstIsBottomPt(bPt1, bPt2)) return outRec1;
            else return outRec2;
        }
        //------------------------------------------------------------------------------

        bool OutRec1RightOfOutRec2(OutRec outRec1, OutRec outRec2)
        {
            do
            {
                outRec1 = outRec1.FirstLeft;
                if (outRec1 == outRec2) return true;
            } while (outRec1 != null);
            return false;
        }
        //------------------------------------------------------------------------------

        private OutRec GetOutRec(int idx)
        {
            OutRec outrec = m_PolyOuts[idx];
            while (outrec != m_PolyOuts[outrec.Idx])
                outrec = m_PolyOuts[outrec.Idx];
            return outrec;
        }
        //------------------------------------------------------------------------------

        private void AppendPolygon(TEdge e1, TEdge e2)
        {
            OutRec outRec1 = m_PolyOuts[e1.OutIdx];
            OutRec outRec2 = m_PolyOuts[e2.OutIdx];

            OutRec holeStateRec;
            if (OutRec1RightOfOutRec2(outRec1, outRec2))
                holeStateRec = outRec2;
            else if (OutRec1RightOfOutRec2(outRec2, outRec1))
                holeStateRec = outRec1;
            else
                holeStateRec = GetLowermostRec(outRec1, outRec2);

            //get the start and ends of both output polygons and
            //join E2 poly onto E1 poly and delete pointers to E2 ...
            OutPt p1_lft = outRec1.Pts;
            OutPt p1_rt = p1_lft.Prev;
            OutPt p2_lft = outRec2.Pts;
            OutPt p2_rt = p2_lft.Prev;

            //join e2 poly onto e1 poly and delete pointers to e2 ...
            if (e1.Side == EdgeSide.esLeft)
            {
                if (e2.Side == EdgeSide.esLeft)
                {
                    //z y x a b c
                    ReversePolyPtLinks(p2_lft);
                    p2_lft.Next = p1_lft;
                    p1_lft.Prev = p2_lft;
                    p1_rt.Next = p2_rt;
                    p2_rt.Prev = p1_rt;
                    outRec1.Pts = p2_rt;
                }
                else
                {
                    //x y z a b c
                    p2_rt.Next = p1_lft;
                    p1_lft.Prev = p2_rt;
                    p2_lft.Prev = p1_rt;
                    p1_rt.Next = p2_lft;
                    outRec1.Pts = p2_lft;
                }
            }
            else
            {
                if (e2.Side == EdgeSide.esRight)
                {
                    //a b c z y x
                    ReversePolyPtLinks(p2_lft);
                    p1_rt.Next = p2_rt;
                    p2_rt.Prev = p1_rt;
                    p2_lft.Next = p1_lft;
                    p1_lft.Prev = p2_lft;
                }
                else
                {
                    //a b c x y z
                    p1_rt.Next = p2_lft;
                    p2_lft.Prev = p1_rt;
                    p1_lft.Prev = p2_rt;
                    p2_rt.Next = p1_lft;
                }
            }

            outRec1.BottomPt = null;
            if (holeStateRec == outRec2)
            {
                if (outRec2.FirstLeft != outRec1)
                    outRec1.FirstLeft = outRec2.FirstLeft;
                outRec1.IsHole = outRec2.IsHole;
            }
            outRec2.Pts = null;
            outRec2.BottomPt = null;

            outRec2.FirstLeft = outRec1;

            int OKIdx = e1.OutIdx;
            int ObsoleteIdx = e2.OutIdx;

            e1.OutIdx = Unassigned; //nb: safe because we only get here via AddLocalMaxPoly
            e2.OutIdx = Unassigned;

            TEdge e = m_ActiveEdges;
            while (e != null)
            {
                if (e.OutIdx == ObsoleteIdx)
                {
                    e.OutIdx = OKIdx;
                    e.Side = e1.Side;
                    break;
                }
                e = e.NextInAEL;
            }
            outRec2.Idx = outRec1.Idx;
        }
        //------------------------------------------------------------------------------

        private void ReversePolyPtLinks(OutPt pp)
        {
            if (pp == null) return;
            OutPt pp1;
            OutPt pp2;
            pp1 = pp;
            do
            {
                pp2 = pp1.Next;
                pp1.Next = pp1.Prev;
                pp1.Prev = pp2;
                pp1 = pp2;
            } while (pp1 != pp);
        }
        //------------------------------------------------------------------------------

        private static void SwapSides(TEdge edge1, TEdge edge2)
        {
            EdgeSide side = edge1.Side;
            edge1.Side = edge2.Side;
            edge2.Side = side;
        }
        //------------------------------------------------------------------------------

        private static void SwapPolyIndexes(TEdge edge1, TEdge edge2)
        {
            int outIdx = edge1.OutIdx;
            edge1.OutIdx = edge2.OutIdx;
            edge2.OutIdx = outIdx;
        }
        //------------------------------------------------------------------------------

        private void IntersectEdges(TEdge e1, TEdge e2, IntPoint pt)
        {
            //e1 will be to the left of e2 BELOW the intersection. Therefore e1 is before
            //e2 in AEL except when e1 is being inserted at the intersection point ...

            bool e1Contributing = (e1.OutIdx >= 0);
            bool e2Contributing = (e2.OutIdx >= 0);


            //if either edge is on an OPEN path ...
            if (e1.WindDelta == 0 || e2.WindDelta == 0)
            {
                //ignore subject-subject open path intersections UNLESS they
                //are both open paths, AND they are both 'contributing maximas' ...
                if (e1.WindDelta == 0 && e2.WindDelta == 0) return;
                //if intersecting a subj line with a subj poly ...
                else if (e1.PolyTyp == e2.PolyTyp &&
                  e1.WindDelta != e2.WindDelta && m_ClipType == ClipType.ctUnion)
                {
                    if (e1.WindDelta == 0)
                    {
                        if (e2Contributing)
                        {
                            AddOutPt(e1, pt);
                            if (e1Contributing) e1.OutIdx = Unassigned;
                        }
                    }
                    else
                    {
                        if (e1Contributing)
                        {
                            AddOutPt(e2, pt);
                            if (e2Contributing) e2.OutIdx = Unassigned;
                        }
                    }
                }
                else if (e1.PolyTyp != e2.PolyTyp)
                {
                    if ((e1.WindDelta == 0) && Math.Abs(e2.WindCnt) == 1 &&
                      (m_ClipType != ClipType.ctUnion || e2.WindCnt2 == 0))
                    {
                        AddOutPt(e1, pt);
                        if (e1Contributing) e1.OutIdx = Unassigned;
                    }
                    else if ((e2.WindDelta == 0) && (Math.Abs(e1.WindCnt) == 1) &&
                      (m_ClipType != ClipType.ctUnion || e1.WindCnt2 == 0))
                    {
                        AddOutPt(e2, pt);
                        if (e2Contributing) e2.OutIdx = Unassigned;
                    }
                }
                return;
            }

            //update winding counts...
            //assumes that e1 will be to the Right of e2 ABOVE the intersection
            if (e1.PolyTyp == e2.PolyTyp)
            {
                if (IsEvenOddFillType(e1))
                {
                    int oldE1WindCnt = e1.WindCnt;
                    e1.WindCnt = e2.WindCnt;
                    e2.WindCnt = oldE1WindCnt;
                }
                else
                {
                    if (e1.WindCnt + e2.WindDelta == 0) e1.WindCnt = -e1.WindCnt;
                    else e1.WindCnt += e2.WindDelta;
                    if (e2.WindCnt - e1.WindDelta == 0) e2.WindCnt = -e2.WindCnt;
                    else e2.WindCnt -= e1.WindDelta;
                }
            }
            else
            {
                if (!IsEvenOddFillType(e2)) e1.WindCnt2 += e2.WindDelta;
                else e1.WindCnt2 = (e1.WindCnt2 == 0) ? 1 : 0;
                if (!IsEvenOddFillType(e1)) e2.WindCnt2 -= e1.WindDelta;
                else e2.WindCnt2 = (e2.WindCnt2 == 0) ? 1 : 0;
            }

            PolyFillType e1FillType, e2FillType, e1FillType2, e2FillType2;
            if (e1.PolyTyp == PolyType.ptSubject)
            {
                e1FillType = m_SubjFillType;
                e1FillType2 = m_ClipFillType;
            }
            else
            {
                e1FillType = m_ClipFillType;
                e1FillType2 = m_SubjFillType;
            }
            if (e2.PolyTyp == PolyType.ptSubject)
            {
                e2FillType = m_SubjFillType;
                e2FillType2 = m_ClipFillType;
            }
            else
            {
                e2FillType = m_ClipFillType;
                e2FillType2 = m_SubjFillType;
            }

            int e1Wc, e2Wc;
            switch (e1FillType)
            {
                case PolyFillType.pftPositive: e1Wc = e1.WindCnt; break;
                case PolyFillType.pftNegative: e1Wc = -e1.WindCnt; break;
                default: e1Wc = Math.Abs(e1.WindCnt); break;
            }
            switch (e2FillType)
            {
                case PolyFillType.pftPositive: e2Wc = e2.WindCnt; break;
                case PolyFillType.pftNegative: e2Wc = -e2.WindCnt; break;
                default: e2Wc = Math.Abs(e2.WindCnt); break;
            }

            if (e1Contributing && e2Contributing)
            {
                if ((e1Wc != 0 && e1Wc != 1) || (e2Wc != 0 && e2Wc != 1) ||
                  (e1.PolyTyp != e2.PolyTyp && m_ClipType != ClipType.ctXor))
                {
                    AddLocalMaxPoly(e1, e2, pt);
                }
                else
                {
                    AddOutPt(e1, pt);
                    AddOutPt(e2, pt);
                    SwapSides(e1, e2);
                    SwapPolyIndexes(e1, e2);
                }
            }
            else if (e1Contributing)
            {
                if (e2Wc == 0 || e2Wc == 1)
                {
                    AddOutPt(e1, pt);
                    SwapSides(e1, e2);
                    SwapPolyIndexes(e1, e2);
                }

            }
            else if (e2Contributing)
            {
                if (e1Wc == 0 || e1Wc == 1)
                {
                    AddOutPt(e2, pt);
                    SwapSides(e1, e2);
                    SwapPolyIndexes(e1, e2);
                }
            }
            else if ((e1Wc == 0 || e1Wc == 1) && (e2Wc == 0 || e2Wc == 1))
            {
                //neither edge is currently contributing ...
                cInt e1Wc2, e2Wc2;
                switch (e1FillType2)
                {
                    case PolyFillType.pftPositive: e1Wc2 = e1.WindCnt2; break;
                    case PolyFillType.pftNegative: e1Wc2 = -e1.WindCnt2; break;
                    default: e1Wc2 = Math.Abs(e1.WindCnt2); break;
                }
                switch (e2FillType2)
                {
                    case PolyFillType.pftPositive: e2Wc2 = e2.WindCnt2; break;
                    case PolyFillType.pftNegative: e2Wc2 = -e2.WindCnt2; break;
                    default: e2Wc2 = Math.Abs(e2.WindCnt2); break;
                }

                if (e1.PolyTyp != e2.PolyTyp)
                {
                    AddLocalMinPoly(e1, e2, pt);
                }
                else if (e1Wc == 1 && e2Wc == 1)
                    switch (m_ClipType)
                    {
                        case ClipType.ctIntersection:
                            if (e1Wc2 > 0 && e2Wc2 > 0)
                                AddLocalMinPoly(e1, e2, pt);
                            break;
                        case ClipType.ctUnion:
                            if (e1Wc2 <= 0 && e2Wc2 <= 0)
                                AddLocalMinPoly(e1, e2, pt);
                            break;
                        case ClipType.ctDifference:
                            if (((e1.PolyTyp == PolyType.ptClip) && (e1Wc2 > 0) && (e2Wc2 > 0)) ||
                                ((e1.PolyTyp == PolyType.ptSubject) && (e1Wc2 <= 0) && (e2Wc2 <= 0)))
                                AddLocalMinPoly(e1, e2, pt);
                            break;
                        case ClipType.ctXor:
                            AddLocalMinPoly(e1, e2, pt);
                            break;
                    }
                else
                    SwapSides(e1, e2);
            }
        }
        //------------------------------------------------------------------------------

        private void DeleteFromSEL(TEdge e)
        {
            TEdge SelPrev = e.PrevInSEL;
            TEdge SelNext = e.NextInSEL;
            if (SelPrev == null && SelNext == null && (e != m_SortedEdges))
                return; //already deleted
            if (SelPrev != null)
                SelPrev.NextInSEL = SelNext;
            else m_SortedEdges = SelNext;
            if (SelNext != null)
                SelNext.PrevInSEL = SelPrev;
            e.NextInSEL = null;
            e.PrevInSEL = null;
        }
        //------------------------------------------------------------------------------

        private void ProcessHorizontals()
        {
            TEdge horzEdge; //m_SortedEdges;
            while (PopEdgeFromSEL(out horzEdge))
                ProcessHorizontal(horzEdge);
        }
        //------------------------------------------------------------------------------

        void GetHorzDirection(TEdge HorzEdge, out Direction Dir, out cInt Left, out cInt Right)
        {
            if (HorzEdge.Bot.X < HorzEdge.Top.X)
            {
                Left = HorzEdge.Bot.X;
                Right = HorzEdge.Top.X;
                Dir = Direction.dLeftToRight;
            }
            else
            {
                Left = HorzEdge.Top.X;
                Right = HorzEdge.Bot.X;
                Dir = Direction.dRightToLeft;
            }
        }
        //------------------------------------------------------------------------

        private void ProcessHorizontal(TEdge horzEdge)
        {
            Direction dir;
            cInt horzLeft, horzRight;
            bool IsOpen = horzEdge.WindDelta == 0;

            GetHorzDirection(horzEdge, out dir, out horzLeft, out horzRight);

            TEdge eLastHorz = horzEdge, eMaxPair = null;
            while (eLastHorz.NextInLML != null && IsHorizontal(eLastHorz.NextInLML))
                eLastHorz = eLastHorz.NextInLML;
            if (eLastHorz.NextInLML == null)
                eMaxPair = GetMaximaPair(eLastHorz);

            Maxima currMax = m_Maxima;
            if (currMax != null)
            {
                //get the first maxima in range (X) ...
                if (dir == Direction.dLeftToRight)
                {
                    while (currMax != null && currMax.X <= horzEdge.Bot.X)
                        currMax = currMax.Next;
                    if (currMax != null && currMax.X >= eLastHorz.Top.X)
                        currMax = null;
                }
                else
                {
                    while (currMax.Next != null && currMax.Next.X < horzEdge.Bot.X)
                        currMax = currMax.Next;
                    if (currMax.X <= eLastHorz.Top.X) currMax = null;
                }
            }

            OutPt op1 = null;
            for (; ; ) //loop through consec. horizontal edges
            {
                bool IsLastHorz = (horzEdge == eLastHorz);
                TEdge e = GetNextInAEL(horzEdge, dir);
                while (e != null)
                {

                    //this code block inserts extra coords into horizontal edges (in output
                    //polygons) whereever maxima touch these horizontal edges. This helps
                    //'simplifying' polygons (ie if the Simplify property is set).
                    if (currMax != null)
                    {
                        if (dir == Direction.dLeftToRight)
                        {
                            while (currMax != null && currMax.X < e.Curr.X)
                            {
                                if (horzEdge.OutIdx >= 0 && !IsOpen)
                                    AddOutPt(horzEdge, new IntPoint(currMax.X, horzEdge.Bot.Y));
                                currMax = currMax.Next;
                            }
                        }
                        else
                        {
                            while (currMax != null && currMax.X > e.Curr.X)
                            {
                                if (horzEdge.OutIdx >= 0 && !IsOpen)
                                    AddOutPt(horzEdge, new IntPoint(currMax.X, horzEdge.Bot.Y));
                                currMax = currMax.Prev;
                            }
                        }
                    }
                    ;

                    if ((dir == Direction.dLeftToRight && e.Curr.X > horzRight) ||
                      (dir == Direction.dRightToLeft && e.Curr.X < horzLeft)) break;

                    //Also break if we've got to the end of an intermediate horizontal edge ...
                    //nb: Smaller Dx's are to the right of larger Dx's ABOVE the horizontal.
                    if (e.Curr.X == horzEdge.Top.X && horzEdge.NextInLML != null &&
                      e.Dx < horzEdge.NextInLML.Dx) break;

                    if (horzEdge.OutIdx >= 0 && !IsOpen)  //note: may be done multiple times
                    {

                        op1 = AddOutPt(horzEdge, e.Curr);
                        TEdge eNextHorz = m_SortedEdges;
                        while (eNextHorz != null)
                        {
                            if (eNextHorz.OutIdx >= 0 &&
                              HorzSegmentsOverlap(horzEdge.Bot.X,
                              horzEdge.Top.X, eNextHorz.Bot.X, eNextHorz.Top.X))
                            {
                                OutPt op2 = GetLastOutPt(eNextHorz);
                                AddJoin(op2, op1, eNextHorz.Top);
                            }
                            eNextHorz = eNextHorz.NextInSEL;
                        }
                        AddGhostJoin(op1, horzEdge.Bot);
                    }

                    //OK, so far we're still in range of the horizontal Edge  but make sure
                    //we're at the last of consec. horizontals when matching with eMaxPair
                    if (e == eMaxPair && IsLastHorz)
                    {
                        if (horzEdge.OutIdx >= 0)
                            AddLocalMaxPoly(horzEdge, eMaxPair, horzEdge.Top);
                        DeleteFromAEL(horzEdge);
                        DeleteFromAEL(eMaxPair);
                        return;
                    }

                    if (dir == Direction.dLeftToRight)
                    {
                        IntPoint Pt = new IntPoint(e.Curr.X, horzEdge.Curr.Y);
                        IntersectEdges(horzEdge, e, Pt);
                    }
                    else
                    {
                        IntPoint Pt = new IntPoint(e.Curr.X, horzEdge.Curr.Y);
                        IntersectEdges(e, horzEdge, Pt);
                    }
                    TEdge eNext = GetNextInAEL(e, dir);
                    SwapPositionsInAEL(horzEdge, e);
                    e = eNext;
                } //end while(e != null)

                //Break out of loop if HorzEdge.NextInLML is not also horizontal ...
                if (horzEdge.NextInLML == null || !IsHorizontal(horzEdge.NextInLML)) break;

                UpdateEdgeIntoAEL(ref horzEdge);
                if (horzEdge.OutIdx >= 0) AddOutPt(horzEdge, horzEdge.Bot);
                GetHorzDirection(horzEdge, out dir, out horzLeft, out horzRight);

            } //end for (;;)

            if (horzEdge.OutIdx >= 0 && op1 == null)
            {
                op1 = GetLastOutPt(horzEdge);
                TEdge eNextHorz = m_SortedEdges;
                while (eNextHorz != null)
                {
                    if (eNextHorz.OutIdx >= 0 &&
                      HorzSegmentsOverlap(horzEdge.Bot.X,
                      horzEdge.Top.X, eNextHorz.Bot.X, eNextHorz.Top.X))
                    {
                        OutPt op2 = GetLastOutPt(eNextHorz);
                        AddJoin(op2, op1, eNextHorz.Top);
                    }
                    eNextHorz = eNextHorz.NextInSEL;
                }
                AddGhostJoin(op1, horzEdge.Top);
            }

            if (horzEdge.NextInLML != null)
            {
                if (horzEdge.OutIdx >= 0)
                {
                    op1 = AddOutPt(horzEdge, horzEdge.Top);

                    UpdateEdgeIntoAEL(ref horzEdge);
                    if (horzEdge.WindDelta == 0) return;
                    //nb: HorzEdge is no longer horizontal here
                    TEdge ePrev = horzEdge.PrevInAEL;
                    TEdge eNext = horzEdge.NextInAEL;
                    if (ePrev != null && ePrev.Curr.X == horzEdge.Bot.X &&
                      ePrev.Curr.Y == horzEdge.Bot.Y && ePrev.WindDelta != 0 &&
                      (ePrev.OutIdx >= 0 && ePrev.Curr.Y > ePrev.Top.Y &&
                      SlopesEqual(horzEdge, ePrev, m_UseFullRange)))
                    {
                        OutPt op2 = AddOutPt(ePrev, horzEdge.Bot);
                        AddJoin(op1, op2, horzEdge.Top);
                    }
                    else if (eNext != null && eNext.Curr.X == horzEdge.Bot.X &&
                      eNext.Curr.Y == horzEdge.Bot.Y && eNext.WindDelta != 0 &&
                      eNext.OutIdx >= 0 && eNext.Curr.Y > eNext.Top.Y &&
                      SlopesEqual(horzEdge, eNext, m_UseFullRange))
                    {
                        OutPt op2 = AddOutPt(eNext, horzEdge.Bot);
                        AddJoin(op1, op2, horzEdge.Top);
                    }
                }
                else
                    UpdateEdgeIntoAEL(ref horzEdge);
            }
            else
            {
                if (horzEdge.OutIdx >= 0) AddOutPt(horzEdge, horzEdge.Top);
                DeleteFromAEL(horzEdge);
            }
        }
        //------------------------------------------------------------------------------

        private TEdge GetNextInAEL(TEdge e, Direction Direction)
        {
            return Direction == Direction.dLeftToRight ? e.NextInAEL : e.PrevInAEL;
        }
        //------------------------------------------------------------------------------

        private bool IsMinima(TEdge e)
        {
            return e != null && (e.Prev.NextInLML != e) && (e.Next.NextInLML != e);
        }
        //------------------------------------------------------------------------------

        private bool IsMaxima(TEdge e, double Y)
        {
            return (e != null && e.Top.Y == Y && e.NextInLML == null);
        }
        //------------------------------------------------------------------------------

        private bool IsIntermediate(TEdge e, double Y)
        {
            return (e.Top.Y == Y && e.NextInLML != null);
        }
        //------------------------------------------------------------------------------

        internal TEdge GetMaximaPair(TEdge e)
        {
            if ((e.Next.Top == e.Top) && e.Next.NextInLML == null)
                return e.Next;
            else if ((e.Prev.Top == e.Top) && e.Prev.NextInLML == null)
                return e.Prev;
            else
                return null;
        }
        //------------------------------------------------------------------------------

        internal TEdge GetMaximaPairEx(TEdge e)
        {
            //as above but returns null if MaxPair isn't in AEL (unless it's horizontal)
            TEdge result = GetMaximaPair(e);
            if (result == null || result.OutIdx == Skip ||
              ((result.NextInAEL == result.PrevInAEL) && !IsHorizontal(result))) return null;
            return result;
        }
        //------------------------------------------------------------------------------

        private bool ProcessIntersections(cInt topY)
        {
            if (m_ActiveEdges == null) return true;
            try
            {
                BuildIntersectList(topY);
                if (m_IntersectList.Count == 0) return true;
                if (m_IntersectList.Count == 1 || FixupIntersectionOrder())
                    ProcessIntersectList();
                else
                    return false;
            }
            catch
            {
                m_SortedEdges = null;
                m_IntersectList.Clear();
                throw new ClipperException("ProcessIntersections error");
            }
            m_SortedEdges = null;
            return true;
        }
        //------------------------------------------------------------------------------

        private void BuildIntersectList(cInt topY)
        {
            if (m_ActiveEdges == null) return;

            //prepare for sorting ...
            TEdge e = m_ActiveEdges;
            m_SortedEdges = e;
            while (e != null)
            {
                e.PrevInSEL = e.PrevInAEL;
                e.NextInSEL = e.NextInAEL;
                e.Curr.X = TopX(e, topY);
                e = e.NextInAEL;
            }

            //bubblesort ...
            bool isModified = true;
            while (isModified && m_SortedEdges != null)
            {
                isModified = false;
                e = m_SortedEdges;
                while (e.NextInSEL != null)
                {
                    TEdge eNext = e.NextInSEL;
                    IntPoint pt;
                    if (e.Curr.X > eNext.Curr.X)
                    {
                        IntersectPoint(e, eNext, out pt);
                        if (pt.Y < topY)
                            pt = new IntPoint(TopX(e, topY), topY);
                        IntersectNode newNode = new IntersectNode();
                        newNode.Edge1 = e;
                        newNode.Edge2 = eNext;
                        newNode.Pt = pt;
                        m_IntersectList.Add(newNode);

                        SwapPositionsInSEL(e, eNext);
                        isModified = true;
                    }
                    else
                        e = eNext;
                }
                if (e.PrevInSEL != null) e.PrevInSEL.NextInSEL = null;
                else break;
            }
            m_SortedEdges = null;
        }
        //------------------------------------------------------------------------------

        private bool EdgesAdjacent(IntersectNode inode)
        {
            return (inode.Edge1.NextInSEL == inode.Edge2) ||
              (inode.Edge1.PrevInSEL == inode.Edge2);
        }
        //------------------------------------------------------------------------------

        private static int IntersectNodeSort(IntersectNode node1, IntersectNode node2)
        {
            //the following typecast is safe because the differences in Pt.Y will
            //be limited to the height of the scanbeam.
            return (int)(node2.Pt.Y - node1.Pt.Y);
        }
        //------------------------------------------------------------------------------

        private bool FixupIntersectionOrder()
        {
            //pre-condition: intersections are sorted bottom-most first.
            //Now it's crucial that intersections are made only between adjacent edges,
            //so to ensure this the order of intersections may need adjusting ...
            m_IntersectList.Sort(m_IntersectNodeComparer);

            CopyAELToSEL();
            int cnt = m_IntersectList.Count;
            for (int i = 0; i < cnt; i++)
            {
                if (!EdgesAdjacent(m_IntersectList[i]))
                {
                    int j = i + 1;
                    while (j < cnt && !EdgesAdjacent(m_IntersectList[j])) j++;
                    if (j == cnt) return false;

                    IntersectNode tmp = m_IntersectList[i];
                    m_IntersectList[i] = m_IntersectList[j];
                    m_IntersectList[j] = tmp;

                }
                SwapPositionsInSEL(m_IntersectList[i].Edge1, m_IntersectList[i].Edge2);
            }
            return true;
        }
        //------------------------------------------------------------------------------

        private void ProcessIntersectList()
        {
            for (int i = 0; i < m_IntersectList.Count; i++)
            {
                IntersectNode iNode = m_IntersectList[i];
                {
                    IntersectEdges(iNode.Edge1, iNode.Edge2, iNode.Pt);
                    SwapPositionsInAEL(iNode.Edge1, iNode.Edge2);
                }
            }
            m_IntersectList.Clear();
        }
        //------------------------------------------------------------------------------

        internal static cInt Round(double value)
        {
            return value < 0 ? (cInt)(value - 0.5) : (cInt)(value + 0.5);
        }
        //------------------------------------------------------------------------------

        private static cInt TopX(TEdge edge, cInt currentY)
        {
            if (currentY == edge.Top.Y)
                return edge.Top.X;
            return edge.Bot.X + Round(edge.Dx * (currentY - edge.Bot.Y));
        }
        //------------------------------------------------------------------------------

        private void IntersectPoint(TEdge edge1, TEdge edge2, out IntPoint ip)
        {
            ip = new IntPoint();
            double b1, b2;
            //nb: with very large coordinate values, it's possible for SlopesEqual() to 
            //return false but for the edge.Dx value be equal due to double precision rounding.
            if (edge1.Dx == edge2.Dx)
            {
                ip.Y = edge1.Curr.Y;
                ip.X = TopX(edge1, ip.Y);
                return;
            }

            if (edge1.Delta.X == 0)
            {
                ip.X = edge1.Bot.X;
                if (IsHorizontal(edge2))
                {
                    ip.Y = edge2.Bot.Y;
                }
                else
                {
                    b2 = edge2.Bot.Y - (edge2.Bot.X / edge2.Dx);
                    ip.Y = Round(ip.X / edge2.Dx + b2);
                }
            }
            else if (edge2.Delta.X == 0)
            {
                ip.X = edge2.Bot.X;
                if (IsHorizontal(edge1))
                {
                    ip.Y = edge1.Bot.Y;
                }
                else
                {
                    b1 = edge1.Bot.Y - (edge1.Bot.X / edge1.Dx);
                    ip.Y = Round(ip.X / edge1.Dx + b1);
                }
            }
            else
            {
                b1 = edge1.Bot.X - edge1.Bot.Y * edge1.Dx;
                b2 = edge2.Bot.X - edge2.Bot.Y * edge2.Dx;
                double q = (b2 - b1) / (edge1.Dx - edge2.Dx);
                ip.Y = Round(q);
                if (Math.Abs(edge1.Dx) < Math.Abs(edge2.Dx))
                    ip.X = Round(edge1.Dx * q + b1);
                else
                    ip.X = Round(edge2.Dx * q + b2);
            }

            if (ip.Y < edge1.Top.Y || ip.Y < edge2.Top.Y)
            {
                if (edge1.Top.Y > edge2.Top.Y)
                    ip.Y = edge1.Top.Y;
                else
                    ip.Y = edge2.Top.Y;
                if (Math.Abs(edge1.Dx) < Math.Abs(edge2.Dx))
                    ip.X = TopX(edge1, ip.Y);
                else
                    ip.X = TopX(edge2, ip.Y);
            }
            //finally, don't allow 'ip' to be BELOW curr.Y (ie bottom of scanbeam) ...
            if (ip.Y > edge1.Curr.Y)
            {
                ip.Y = edge1.Curr.Y;
                //better to use the more vertical edge to derive X ...
                if (Math.Abs(edge1.Dx) > Math.Abs(edge2.Dx))
                    ip.X = TopX(edge2, ip.Y);
                else
                    ip.X = TopX(edge1, ip.Y);
            }
        }
        //------------------------------------------------------------------------------

        private void ProcessEdgesAtTopOfScanbeam(cInt topY)
        {
            TEdge e = m_ActiveEdges;
            while (e != null)
            {
                //1. process maxima, treating them as if they're 'bent' horizontal edges,
                //   but exclude maxima with horizontal edges. nb: e can't be a horizontal.
                bool IsMaximaEdge = IsMaxima(e, topY);

                if (IsMaximaEdge)
                {
                    TEdge eMaxPair = GetMaximaPairEx(e);
                    IsMaximaEdge = (eMaxPair == null || !IsHorizontal(eMaxPair));
                }

                if (IsMaximaEdge)
                {
                    if (StrictlySimple) InsertMaxima(e.Top.X);
                    TEdge ePrev = e.PrevInAEL;
                    DoMaxima(e);
                    if (ePrev == null) e = m_ActiveEdges;
                    else e = ePrev.NextInAEL;
                }
                else
                {
                    //2. promote horizontal edges, otherwise update Curr.X and Curr.Y ...
                    if (IsIntermediate(e, topY) && IsHorizontal(e.NextInLML))
                    {
                        UpdateEdgeIntoAEL(ref e);
                        if (e.OutIdx >= 0)
                            AddOutPt(e, e.Bot);
                        AddEdgeToSEL(e);
                    }
                    else
                    {
                        e.Curr.X = TopX(e, topY);
                        e.Curr.Y = topY;
                    }
                    //When StrictlySimple and 'e' is being touched by another edge, then
                    //make sure both edges have a vertex here ...
                    if (StrictlySimple)
                    {
                        TEdge ePrev = e.PrevInAEL;
                        if ((e.OutIdx >= 0) && (e.WindDelta != 0) && ePrev != null &&
                          (ePrev.OutIdx >= 0) && (ePrev.Curr.X == e.Curr.X) &&
                          (ePrev.WindDelta != 0))
                        {
                            IntPoint ip = new IntPoint(e.Curr);
                            OutPt op = AddOutPt(ePrev, ip);
                            OutPt op2 = AddOutPt(e, ip);
                            AddJoin(op, op2, ip); //StrictlySimple (type-3) join
                        }
                    }

                    e = e.NextInAEL;
                }
            }

            //3. Process horizontals at the Top of the scanbeam ...
            ProcessHorizontals();
            m_Maxima = null;

            //4. Promote intermediate vertices ...
            e = m_ActiveEdges;
            while (e != null)
            {
                if (IsIntermediate(e, topY))
                {
                    OutPt op = null;
                    if (e.OutIdx >= 0)
                        op = AddOutPt(e, e.Top);
                    UpdateEdgeIntoAEL(ref e);

                    //if output polygons share an edge, they'll need joining later ...
                    TEdge ePrev = e.PrevInAEL;
                    TEdge eNext = e.NextInAEL;
                    if (ePrev != null && ePrev.Curr.X == e.Bot.X &&
                      ePrev.Curr.Y == e.Bot.Y && op != null &&
                      ePrev.OutIdx >= 0 && ePrev.Curr.Y > ePrev.Top.Y &&
                      SlopesEqual(e.Curr, e.Top, ePrev.Curr, ePrev.Top, m_UseFullRange) &&
                      (e.WindDelta != 0) && (ePrev.WindDelta != 0))
                    {
                        OutPt op2 = AddOutPt(ePrev, e.Bot);
                        AddJoin(op, op2, e.Top);
                    }
                    else if (eNext != null && eNext.Curr.X == e.Bot.X &&
                      eNext.Curr.Y == e.Bot.Y && op != null &&
                      eNext.OutIdx >= 0 && eNext.Curr.Y > eNext.Top.Y &&
                      SlopesEqual(e.Curr, e.Top, eNext.Curr, eNext.Top, m_UseFullRange) &&
                      (e.WindDelta != 0) && (eNext.WindDelta != 0))
                    {
                        OutPt op2 = AddOutPt(eNext, e.Bot);
                        AddJoin(op, op2, e.Top);
                    }
                }
                e = e.NextInAEL;
            }
        }
        //------------------------------------------------------------------------------

        private void DoMaxima(TEdge e)
        {
            TEdge eMaxPair = GetMaximaPairEx(e);
            if (eMaxPair == null)
            {
                if (e.OutIdx >= 0)
                    AddOutPt(e, e.Top);
                DeleteFromAEL(e);
                return;
            }

            TEdge eNext = e.NextInAEL;
            while (eNext != null && eNext != eMaxPair)
            {
                IntersectEdges(e, eNext, e.Top);
                SwapPositionsInAEL(e, eNext);
                eNext = e.NextInAEL;
            }

            if (e.OutIdx == Unassigned && eMaxPair.OutIdx == Unassigned)
            {
                DeleteFromAEL(e);
                DeleteFromAEL(eMaxPair);
            }
            else if (e.OutIdx >= 0 && eMaxPair.OutIdx >= 0)
            {
                if (e.OutIdx >= 0) AddLocalMaxPoly(e, eMaxPair, e.Top);
                DeleteFromAEL(e);
                DeleteFromAEL(eMaxPair);
            }
            else if (e.WindDelta == 0)
            {
                if (e.OutIdx >= 0)
                {
                    AddOutPt(e, e.Top);
                    e.OutIdx = Unassigned;
                }
                DeleteFromAEL(e);

                if (eMaxPair.OutIdx >= 0)
                {
                    AddOutPt(eMaxPair, e.Top);
                    eMaxPair.OutIdx = Unassigned;
                }
                DeleteFromAEL(eMaxPair);
            }
            else throw new ClipperException("DoMaxima error");
        }
        //------------------------------------------------------------------------------

        public static void ReversePaths(Paths polys)
        {
            foreach (var poly in polys) { poly.Reverse(); }
        }
        //------------------------------------------------------------------------------

        public static bool Orientation(Path poly)
        {
            return Area(poly) >= 0;
        }
        //------------------------------------------------------------------------------

        private int PointCount(OutPt pts)
        {
            if (pts == null) return 0;
            int result = 0;
            OutPt p = pts;
            do
            {
                result++;
                p = p.Next;
            }
            while (p != pts);
            return result;
        }
        //------------------------------------------------------------------------------

        private void BuildResult(Paths polyg)
        {
            polyg.Clear();
            polyg.Capacity = m_PolyOuts.Count;
            for (int i = 0; i < m_PolyOuts.Count; i++)
            {
                OutRec outRec = m_PolyOuts[i];
                if (outRec.Pts == null) continue;
                OutPt p = outRec.Pts.Prev;
                int cnt = PointCount(p);
                if (cnt < 2) continue;
                Path pg = new Path(cnt);
                for (int j = 0; j < cnt; j++)
                {
                    pg.Add(p.Pt);
                    p = p.Prev;
                }
                polyg.Add(pg);
            }
        }
        //------------------------------------------------------------------------------

        private void BuildResult2(PolyTree polytree)
        {
            polytree.Clear();

            //add each output polygon/contour to polytree ...
            polytree.m_AllPolys.Capacity = m_PolyOuts.Count;
            for (int i = 0; i < m_PolyOuts.Count; i++)
            {
                OutRec outRec = m_PolyOuts[i];
                int cnt = PointCount(outRec.Pts);
                if ((outRec.IsOpen && cnt < 2) ||
                  (!outRec.IsOpen && cnt < 3)) continue;
                FixHoleLinkage(outRec);
                PolyNode pn = new PolyNode();
                polytree.m_AllPolys.Add(pn);
                outRec.PolyNode = pn;
                pn.m_polygon.Capacity = cnt;
                OutPt op = outRec.Pts.Prev;
                for (int j = 0; j < cnt; j++)
                {
                    pn.m_polygon.Add(op.Pt);
                    op = op.Prev;
                }
            }

            //fixup PolyNode links etc ...
            polytree.m_Childs.Capacity = m_PolyOuts.Count;
            for (int i = 0; i < m_PolyOuts.Count; i++)
            {
                OutRec outRec = m_PolyOuts[i];
                if (outRec.PolyNode == null) continue;
                else if (outRec.IsOpen)
                {
                    outRec.PolyNode.IsOpen = true;
                    polytree.AddChild(outRec.PolyNode);
                }
                else if (outRec.FirstLeft != null &&
                  outRec.FirstLeft.PolyNode != null)
                    outRec.FirstLeft.PolyNode.AddChild(outRec.PolyNode);
                else
                    polytree.AddChild(outRec.PolyNode);
            }
        }
        //------------------------------------------------------------------------------

        private void FixupOutPolyline(OutRec outrec)
        {
            OutPt pp = outrec.Pts;
            OutPt lastPP = pp.Prev;
            while (pp != lastPP)
            {
                pp = pp.Next;
                if (pp.Pt == pp.Prev.Pt)
                {
                    if (pp == lastPP) lastPP = pp.Prev;
                    OutPt tmpPP = pp.Prev;
                    tmpPP.Next = pp.Next;
                    pp.Next.Prev = tmpPP;
                    pp = tmpPP;
                }
            }
            if (pp == pp.Prev) outrec.Pts = null;
        }
        //------------------------------------------------------------------------------

        private void FixupOutPolygon(OutRec outRec)
        {
            //FixupOutPolygon() - removes duplicate points and simplifies consecutive
            //parallel edges by removing the middle vertex.
            OutPt lastOK = null;
            outRec.BottomPt = null;
            OutPt pp = outRec.Pts;
            bool preserveCol = PreserveCollinear || StrictlySimple;
            for (; ; )
            {
                if (pp.Prev == pp || pp.Prev == pp.Next)
                {
                    outRec.Pts = null;
                    return;
                }
                //test for duplicate points and collinear edges ...
                if ((pp.Pt == pp.Next.Pt) || (pp.Pt == pp.Prev.Pt) ||
                  (SlopesEqual(pp.Prev.Pt, pp.Pt, pp.Next.Pt, m_UseFullRange) &&
                  (!preserveCol || !Pt2IsBetweenPt1AndPt3(pp.Prev.Pt, pp.Pt, pp.Next.Pt))))
                {
                    lastOK = null;
                    pp.Prev.Next = pp.Next;
                    pp.Next.Prev = pp.Prev;
                    pp = pp.Prev;
                }
                else if (pp == lastOK) break;
                else
                {
                    if (lastOK == null) lastOK = pp;
                    pp = pp.Next;
                }
            }
            outRec.Pts = pp;
        }
        //------------------------------------------------------------------------------

        OutPt DupOutPt(OutPt outPt, bool InsertAfter)
        {
            OutPt result = new OutPt();
            result.Pt = outPt.Pt;
            result.Idx = outPt.Idx;
            if (InsertAfter)
            {
                result.Next = outPt.Next;
                result.Prev = outPt;
                outPt.Next.Prev = result;
                outPt.Next = result;
            }
            else
            {
                result.Prev = outPt.Prev;
                result.Next = outPt;
                outPt.Prev.Next = result;
                outPt.Prev = result;
            }
            return result;
        }
        //------------------------------------------------------------------------------

        bool GetOverlap(cInt a1, cInt a2, cInt b1, cInt b2, out cInt Left, out cInt Right)
        {
            if (a1 < a2)
            {
                if (b1 < b2) { Left = Math.Max(a1, b1); Right = Math.Min(a2, b2); }
                else { Left = Math.Max(a1, b2); Right = Math.Min(a2, b1); }
            }
            else
            {
                if (b1 < b2) { Left = Math.Max(a2, b1); Right = Math.Min(a1, b2); }
                else { Left = Math.Max(a2, b2); Right = Math.Min(a1, b1); }
            }
            return Left < Right;
        }
        //------------------------------------------------------------------------------

        bool JoinHorz(OutPt op1, OutPt op1b, OutPt op2, OutPt op2b,
          IntPoint Pt, bool DiscardLeft)
        {
            Direction Dir1 = (op1.Pt.X > op1b.Pt.X ?
              Direction.dRightToLeft : Direction.dLeftToRight);
            Direction Dir2 = (op2.Pt.X > op2b.Pt.X ?
              Direction.dRightToLeft : Direction.dLeftToRight);
            if (Dir1 == Dir2) return false;

            //When DiscardLeft, we want Op1b to be on the Left of Op1, otherwise we
            //want Op1b to be on the Right. (And likewise with Op2 and Op2b.)
            //So, to facilitate this while inserting Op1b and Op2b ...
            //when DiscardLeft, make sure we're AT or RIGHT of Pt before adding Op1b,
            //otherwise make sure we're AT or LEFT of Pt. (Likewise with Op2b.)
            if (Dir1 == Direction.dLeftToRight)
            {
                while (op1.Next.Pt.X <= Pt.X &&
                  op1.Next.Pt.X >= op1.Pt.X && op1.Next.Pt.Y == Pt.Y)
                    op1 = op1.Next;
                if (DiscardLeft && (op1.Pt.X != Pt.X)) op1 = op1.Next;
                op1b = DupOutPt(op1, !DiscardLeft);
                if (op1b.Pt != Pt)
                {
                    op1 = op1b;
                    op1.Pt = Pt;
                    op1b = DupOutPt(op1, !DiscardLeft);
                }
            }
            else
            {
                while (op1.Next.Pt.X >= Pt.X &&
                  op1.Next.Pt.X <= op1.Pt.X && op1.Next.Pt.Y == Pt.Y)
                    op1 = op1.Next;
                if (!DiscardLeft && (op1.Pt.X != Pt.X)) op1 = op1.Next;
                op1b = DupOutPt(op1, DiscardLeft);
                if (op1b.Pt != Pt)
                {
                    op1 = op1b;
                    op1.Pt = Pt;
                    op1b = DupOutPt(op1, DiscardLeft);
                }
            }

            if (Dir2 == Direction.dLeftToRight)
            {
                while (op2.Next.Pt.X <= Pt.X &&
                  op2.Next.Pt.X >= op2.Pt.X && op2.Next.Pt.Y == Pt.Y)
                    op2 = op2.Next;
                if (DiscardLeft && (op2.Pt.X != Pt.X)) op2 = op2.Next;
                op2b = DupOutPt(op2, !DiscardLeft);
                if (op2b.Pt != Pt)
                {
                    op2 = op2b;
                    op2.Pt = Pt;
                    op2b = DupOutPt(op2, !DiscardLeft);
                }
                ;
            }
            else
            {
                while (op2.Next.Pt.X >= Pt.X &&
                  op2.Next.Pt.X <= op2.Pt.X && op2.Next.Pt.Y == Pt.Y)
                    op2 = op2.Next;
                if (!DiscardLeft && (op2.Pt.X != Pt.X)) op2 = op2.Next;
                op2b = DupOutPt(op2, DiscardLeft);
                if (op2b.Pt != Pt)
                {
                    op2 = op2b;
                    op2.Pt = Pt;
                    op2b = DupOutPt(op2, DiscardLeft);
                }
                ;
            }
            ;

            if ((Dir1 == Direction.dLeftToRight) == DiscardLeft)
            {
                op1.Prev = op2;
                op2.Next = op1;
                op1b.Next = op2b;
                op2b.Prev = op1b;
            }
            else
            {
                op1.Next = op2;
                op2.Prev = op1;
                op1b.Prev = op2b;
                op2b.Next = op1b;
            }
            return true;
        }
        //------------------------------------------------------------------------------

        private bool JoinPoints(Join j, OutRec outRec1, OutRec outRec2)
        {
            OutPt op1 = j.OutPt1, op1b;
            OutPt op2 = j.OutPt2, op2b;

            //There are 3 kinds of joins for output polygons ...
            //1. Horizontal joins where Join.OutPt1 & Join.OutPt2 are vertices anywhere
            //along (horizontal) collinear edges (& Join.OffPt is on the same horizontal).
            //2. Non-horizontal joins where Join.OutPt1 & Join.OutPt2 are at the same
            //location at the Bottom of the overlapping segment (& Join.OffPt is above).
            //3. StrictlySimple joins where edges touch but are not collinear and where
            //Join.OutPt1, Join.OutPt2 & Join.OffPt all share the same point.
            bool isHorizontal = (j.OutPt1.Pt.Y == j.OffPt.Y);

            if (isHorizontal && (j.OffPt == j.OutPt1.Pt) && (j.OffPt == j.OutPt2.Pt))
            {
                //Strictly Simple join ...
                if (outRec1 != outRec2) return false;
                op1b = j.OutPt1.Next;
                while (op1b != op1 && (op1b.Pt == j.OffPt))
                    op1b = op1b.Next;
                bool reverse1 = (op1b.Pt.Y > j.OffPt.Y);
                op2b = j.OutPt2.Next;
                while (op2b != op2 && (op2b.Pt == j.OffPt))
                    op2b = op2b.Next;
                bool reverse2 = (op2b.Pt.Y > j.OffPt.Y);
                if (reverse1 == reverse2) return false;
                if (reverse1)
                {
                    op1b = DupOutPt(op1, false);
                    op2b = DupOutPt(op2, true);
                    op1.Prev = op2;
                    op2.Next = op1;
                    op1b.Next = op2b;
                    op2b.Prev = op1b;
                    j.OutPt1 = op1;
                    j.OutPt2 = op1b;
                    return true;
                }
                else
                {
                    op1b = DupOutPt(op1, true);
                    op2b = DupOutPt(op2, false);
                    op1.Next = op2;
                    op2.Prev = op1;
                    op1b.Prev = op2b;
                    op2b.Next = op1b;
                    j.OutPt1 = op1;
                    j.OutPt2 = op1b;
                    return true;
                }
            }
            else if (isHorizontal)
            {
                //treat horizontal joins differently to non-horizontal joins since with
                //them we're not yet sure where the overlapping is. OutPt1.Pt & OutPt2.Pt
                //may be anywhere along the horizontal edge.
                op1b = op1;
                while (op1.Prev.Pt.Y == op1.Pt.Y && op1.Prev != op1b && op1.Prev != op2)
                    op1 = op1.Prev;
                while (op1b.Next.Pt.Y == op1b.Pt.Y && op1b.Next != op1 && op1b.Next != op2)
                    op1b = op1b.Next;
                if (op1b.Next == op1 || op1b.Next == op2) return false; //a flat 'polygon'

                op2b = op2;
                while (op2.Prev.Pt.Y == op2.Pt.Y && op2.Prev != op2b && op2.Prev != op1b)
                    op2 = op2.Prev;
                while (op2b.Next.Pt.Y == op2b.Pt.Y && op2b.Next != op2 && op2b.Next != op1)
                    op2b = op2b.Next;
                if (op2b.Next == op2 || op2b.Next == op1) return false; //a flat 'polygon'

                cInt Left, Right;
                //Op1 -. Op1b & Op2 -. Op2b are the extremites of the horizontal edges
                if (!GetOverlap(op1.Pt.X, op1b.Pt.X, op2.Pt.X, op2b.Pt.X, out Left, out Right))
                    return false;

                //DiscardLeftSide: when overlapping edges are joined, a spike will created
                //which needs to be cleaned up. However, we don't want Op1 or Op2 caught up
                //on the discard Side as either may still be needed for other joins ...
                IntPoint Pt;
                bool DiscardLeftSide;
                if (op1.Pt.X >= Left && op1.Pt.X <= Right)
                {
                    Pt = op1.Pt; DiscardLeftSide = (op1.Pt.X > op1b.Pt.X);
                }
                else if (op2.Pt.X >= Left && op2.Pt.X <= Right)
                {
                    Pt = op2.Pt; DiscardLeftSide = (op2.Pt.X > op2b.Pt.X);
                }
                else if (op1b.Pt.X >= Left && op1b.Pt.X <= Right)
                {
                    Pt = op1b.Pt; DiscardLeftSide = op1b.Pt.X > op1.Pt.X;
                }
                else
                {
                    Pt = op2b.Pt; DiscardLeftSide = (op2b.Pt.X > op2.Pt.X);
                }
                j.OutPt1 = op1;
                j.OutPt2 = op2;
                return JoinHorz(op1, op1b, op2, op2b, Pt, DiscardLeftSide);
            }
            else
            {
                //nb: For non-horizontal joins ...
                //    1. Jr.OutPt1.Pt.Y == Jr.OutPt2.Pt.Y
                //    2. Jr.OutPt1.Pt > Jr.OffPt.Y

                //make sure the polygons are correctly oriented ...
                op1b = op1.Next;
                while ((op1b.Pt == op1.Pt) && (op1b != op1)) op1b = op1b.Next;
                bool Reverse1 = ((op1b.Pt.Y > op1.Pt.Y) ||
                  !SlopesEqual(op1.Pt, op1b.Pt, j.OffPt, m_UseFullRange));
                if (Reverse1)
                {
                    op1b = op1.Prev;
                    while ((op1b.Pt == op1.Pt) && (op1b != op1)) op1b = op1b.Prev;
                    if ((op1b.Pt.Y > op1.Pt.Y) ||
                      !SlopesEqual(op1.Pt, op1b.Pt, j.OffPt, m_UseFullRange)) return false;
                }
                ;
                op2b = op2.Next;
                while ((op2b.Pt == op2.Pt) && (op2b != op2)) op2b = op2b.Next;
                bool Reverse2 = ((op2b.Pt.Y > op2.Pt.Y) ||
                  !SlopesEqual(op2.Pt, op2b.Pt, j.OffPt, m_UseFullRange));
                if (Reverse2)
                {
                    op2b = op2.Prev;
                    while ((op2b.Pt == op2.Pt) && (op2b != op2)) op2b = op2b.Prev;
                    if ((op2b.Pt.Y > op2.Pt.Y) ||
                      !SlopesEqual(op2.Pt, op2b.Pt, j.OffPt, m_UseFullRange)) return false;
                }

                if ((op1b == op1) || (op2b == op2) || (op1b == op2b) ||
                  ((outRec1 == outRec2) && (Reverse1 == Reverse2))) return false;

                if (Reverse1)
                {
                    op1b = DupOutPt(op1, false);
                    op2b = DupOutPt(op2, true);
                    op1.Prev = op2;
                    op2.Next = op1;
                    op1b.Next = op2b;
                    op2b.Prev = op1b;
                    j.OutPt1 = op1;
                    j.OutPt2 = op1b;
                    return true;
                }
                else
                {
                    op1b = DupOutPt(op1, true);
                    op2b = DupOutPt(op2, false);
                    op1.Next = op2;
                    op2.Prev = op1;
                    op1b.Prev = op2b;
                    op2b.Next = op1b;
                    j.OutPt1 = op1;
                    j.OutPt2 = op1b;
                    return true;
                }
            }
        }
        //----------------------------------------------------------------------

        public static int PointInPolygon(IntPoint pt, Path path)
        {
            //returns 0 if false, +1 if true, -1 if pt ON polygon boundary
            //See "The Point in Polygon Problem for Arbitrary Polygons" by Hormann & Agathos
            //http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.88.5498&rep=rep1&type=pdf
            int result = 0, cnt = path.Count;
            if (cnt < 3) return 0;
            IntPoint ip = path[0];
            for (int i = 1; i <= cnt; ++i)
            {
                IntPoint ipNext = (i == cnt ? path[0] : path[i]);
                if (ipNext.Y == pt.Y)
                {
                    if ((ipNext.X == pt.X) || (ip.Y == pt.Y &&
                      ((ipNext.X > pt.X) == (ip.X < pt.X)))) return -1;
                }
                if ((ip.Y < pt.Y) != (ipNext.Y < pt.Y))
                {
                    if (ip.X >= pt.X)
                    {
                        if (ipNext.X > pt.X) result = 1 - result;
                        else
                        {
                            double d = (double)(ip.X - pt.X) * (ipNext.Y - pt.Y) -
                              (double)(ipNext.X - pt.X) * (ip.Y - pt.Y);
                            if (d == 0) return -1;
                            else if ((d > 0) == (ipNext.Y > ip.Y)) result = 1 - result;
                        }
                    }
                    else
                    {
                        if (ipNext.X > pt.X)
                        {
                            double d = (double)(ip.X - pt.X) * (ipNext.Y - pt.Y) -
                              (double)(ipNext.X - pt.X) * (ip.Y - pt.Y);
                            if (d == 0) return -1;
                            else if ((d > 0) == (ipNext.Y > ip.Y)) result = 1 - result;
                        }
                    }
                }
                ip = ipNext;
            }
            return result;
        }
        //------------------------------------------------------------------------------

        //See "The Point in Polygon Problem for Arbitrary Polygons" by Hormann & Agathos
        //http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.88.5498&rep=rep1&type=pdf
        private static int PointInPolygon(IntPoint pt, OutPt op)
        {
            //returns 0 if false, +1 if true, -1 if pt ON polygon boundary
            int result = 0;
            OutPt startOp = op;
            cInt ptx = pt.X, pty = pt.Y;
            cInt poly0x = op.Pt.X, poly0y = op.Pt.Y;
            do
            {
                op = op.Next;
                cInt poly1x = op.Pt.X, poly1y = op.Pt.Y;

                if (poly1y == pty)
                {
                    if ((poly1x == ptx) || (poly0y == pty &&
                      ((poly1x > ptx) == (poly0x < ptx)))) return -1;
                }
                if ((poly0y < pty) != (poly1y < pty))
                {
                    if (poly0x >= ptx)
                    {
                        if (poly1x > ptx) result = 1 - result;
                        else
                        {
                            double d = (double)(poly0x - ptx) * (poly1y - pty) -
                              (double)(poly1x - ptx) * (poly0y - pty);
                            if (d == 0) return -1;
                            if ((d > 0) == (poly1y > poly0y)) result = 1 - result;
                        }
                    }
                    else
                    {
                        if (poly1x > ptx)
                        {
                            double d = (double)(poly0x - ptx) * (poly1y - pty) -
                              (double)(poly1x - ptx) * (poly0y - pty);
                            if (d == 0) return -1;
                            if ((d > 0) == (poly1y > poly0y)) result = 1 - result;
                        }
                    }
                }
                poly0x = poly1x; poly0y = poly1y;
            } while (startOp != op);
            return result;
        }
        //------------------------------------------------------------------------------

        private static bool Poly2ContainsPoly1(OutPt outPt1, OutPt outPt2)
        {
            OutPt op = outPt1;
            do
            {
                //nb: PointInPolygon returns 0 if false, +1 if true, -1 if pt on polygon
                int res = PointInPolygon(op.Pt, outPt2);
                if (res >= 0) return res > 0;
                op = op.Next;
            }
            while (op != outPt1);
            return true;
        }
        //----------------------------------------------------------------------

        private void FixupFirstLefts1(OutRec OldOutRec, OutRec NewOutRec)
        {
            foreach (OutRec outRec in m_PolyOuts)
            {
                OutRec firstLeft = ParseFirstLeft(outRec.FirstLeft);
                if (outRec.Pts != null && firstLeft == OldOutRec)
                {
                    if (Poly2ContainsPoly1(outRec.Pts, NewOutRec.Pts))
                        outRec.FirstLeft = NewOutRec;
                }
            }
        }
        //----------------------------------------------------------------------

        private void FixupFirstLefts2(OutRec innerOutRec, OutRec outerOutRec)
        {
            //A polygon has split into two such that one is now the inner of the other.
            //It's possible that these polygons now wrap around other polygons, so check
            //every polygon that's also contained by OuterOutRec's FirstLeft container
            //(including nil) to see if they've become inner to the new inner polygon ...
            OutRec orfl = outerOutRec.FirstLeft;
            foreach (OutRec outRec in m_PolyOuts)
            {
                if (outRec.Pts == null || outRec == outerOutRec || outRec == innerOutRec)
                    continue;
                OutRec firstLeft = ParseFirstLeft(outRec.FirstLeft);
                if (firstLeft != orfl && firstLeft != innerOutRec && firstLeft != outerOutRec)
                    continue;
                if (Poly2ContainsPoly1(outRec.Pts, innerOutRec.Pts))
                    outRec.FirstLeft = innerOutRec;
                else if (Poly2ContainsPoly1(outRec.Pts, outerOutRec.Pts))
                    outRec.FirstLeft = outerOutRec;
                else if (outRec.FirstLeft == innerOutRec || outRec.FirstLeft == outerOutRec)
                    outRec.FirstLeft = orfl;
            }
        }
        //----------------------------------------------------------------------

        private void FixupFirstLefts3(OutRec OldOutRec, OutRec NewOutRec)
        {
            //same as FixupFirstLefts1 but doesn't call Poly2ContainsPoly1()
            foreach (OutRec outRec in m_PolyOuts)
            {
                OutRec firstLeft = ParseFirstLeft(outRec.FirstLeft);
                if (outRec.Pts != null && firstLeft == OldOutRec)
                    outRec.FirstLeft = NewOutRec;
            }
        }
        //----------------------------------------------------------------------

        private static OutRec ParseFirstLeft(OutRec FirstLeft)
        {
            while (FirstLeft != null && FirstLeft.Pts == null)
                FirstLeft = FirstLeft.FirstLeft;
            return FirstLeft;
        }
        //------------------------------------------------------------------------------

        private void JoinCommonEdges()
        {
            for (int i = 0; i < m_Joins.Count; i++)
            {
                Join join = m_Joins[i];

                OutRec outRec1 = GetOutRec(join.OutPt1.Idx);
                OutRec outRec2 = GetOutRec(join.OutPt2.Idx);

                if (outRec1.Pts == null || outRec2.Pts == null) continue;
                if (outRec1.IsOpen || outRec2.IsOpen) continue;

                //get the polygon fragment with the correct hole state (FirstLeft)
                //before calling JoinPoints() ...
                OutRec holeStateRec;
                if (outRec1 == outRec2) holeStateRec = outRec1;
                else if (OutRec1RightOfOutRec2(outRec1, outRec2)) holeStateRec = outRec2;
                else if (OutRec1RightOfOutRec2(outRec2, outRec1)) holeStateRec = outRec1;
                else holeStateRec = GetLowermostRec(outRec1, outRec2);

                if (!JoinPoints(join, outRec1, outRec2)) continue;

                if (outRec1 == outRec2)
                {
                    //instead of joining two polygons, we've just created a new one by
                    //splitting one polygon into two.
                    outRec1.Pts = join.OutPt1;
                    outRec1.BottomPt = null;
                    outRec2 = CreateOutRec();
                    outRec2.Pts = join.OutPt2;

                    //update all OutRec2.Pts Idx's ...
                    UpdateOutPtIdxs(outRec2);

                    if (Poly2ContainsPoly1(outRec2.Pts, outRec1.Pts))
                    {
                        //outRec1 contains outRec2 ...
                        outRec2.IsHole = !outRec1.IsHole;
                        outRec2.FirstLeft = outRec1;

                        if (m_UsingPolyTree) FixupFirstLefts2(outRec2, outRec1);

                        if ((outRec2.IsHole ^ ReverseSolution) == (Area(outRec2) > 0))
                            ReversePolyPtLinks(outRec2.Pts);

                    }
                    else if (Poly2ContainsPoly1(outRec1.Pts, outRec2.Pts))
                    {
                        //outRec2 contains outRec1 ...
                        outRec2.IsHole = outRec1.IsHole;
                        outRec1.IsHole = !outRec2.IsHole;
                        outRec2.FirstLeft = outRec1.FirstLeft;
                        outRec1.FirstLeft = outRec2;

                        if (m_UsingPolyTree) FixupFirstLefts2(outRec1, outRec2);

                        if ((outRec1.IsHole ^ ReverseSolution) == (Area(outRec1) > 0))
                            ReversePolyPtLinks(outRec1.Pts);
                    }
                    else
                    {
                        //the 2 polygons are completely separate ...
                        outRec2.IsHole = outRec1.IsHole;
                        outRec2.FirstLeft = outRec1.FirstLeft;

                        //fixup FirstLeft pointers that may need reassigning to OutRec2
                        if (m_UsingPolyTree) FixupFirstLefts1(outRec1, outRec2);
                    }

                }
                else
                {
                    //joined 2 polygons together ...

                    outRec2.Pts = null;
                    outRec2.BottomPt = null;
                    outRec2.Idx = outRec1.Idx;

                    outRec1.IsHole = holeStateRec.IsHole;
                    if (holeStateRec == outRec2)
                        outRec1.FirstLeft = outRec2.FirstLeft;
                    outRec2.FirstLeft = outRec1;

                    //fixup FirstLeft pointers that may need reassigning to OutRec1
                    if (m_UsingPolyTree) FixupFirstLefts3(outRec2, outRec1);
                }
            }
        }
        //------------------------------------------------------------------------------

        private void UpdateOutPtIdxs(OutRec outrec)
        {
            OutPt op = outrec.Pts;
            do
            {
                op.Idx = outrec.Idx;
                op = op.Prev;
            }
            while (op != outrec.Pts);
        }
        //------------------------------------------------------------------------------

        private void DoSimplePolygons()
        {
            int i = 0;
            while (i < m_PolyOuts.Count)
            {
                OutRec outrec = m_PolyOuts[i++];
                OutPt op = outrec.Pts;
                if (op == null || outrec.IsOpen) continue;
                do //for each Pt in Polygon until duplicate found do ...
                {
                    OutPt op2 = op.Next;
                    while (op2 != outrec.Pts)
                    {
                        if ((op.Pt == op2.Pt) && op2.Next != op && op2.Prev != op)
                        {
                            //split the polygon into two ...
                            OutPt op3 = op.Prev;
                            OutPt op4 = op2.Prev;
                            op.Prev = op4;
                            op4.Next = op;
                            op2.Prev = op3;
                            op3.Next = op2;

                            outrec.Pts = op;
                            OutRec outrec2 = CreateOutRec();
                            outrec2.Pts = op2;
                            UpdateOutPtIdxs(outrec2);
                            if (Poly2ContainsPoly1(outrec2.Pts, outrec.Pts))
                            {
                                //OutRec2 is contained by OutRec1 ...
                                outrec2.IsHole = !outrec.IsHole;
                                outrec2.FirstLeft = outrec;
                                if (m_UsingPolyTree) FixupFirstLefts2(outrec2, outrec);
                            }
                            else
                              if (Poly2ContainsPoly1(outrec.Pts, outrec2.Pts))
                            {
                                //OutRec1 is contained by OutRec2 ...
                                outrec2.IsHole = outrec.IsHole;
                                outrec.IsHole = !outrec2.IsHole;
                                outrec2.FirstLeft = outrec.FirstLeft;
                                outrec.FirstLeft = outrec2;
                                if (m_UsingPolyTree) FixupFirstLefts2(outrec, outrec2);
                            }
                            else
                            {
                                //the 2 polygons are separate ...
                                outrec2.IsHole = outrec.IsHole;
                                outrec2.FirstLeft = outrec.FirstLeft;
                                if (m_UsingPolyTree) FixupFirstLefts1(outrec, outrec2);
                            }
                            op2 = op; //ie get ready for the next iteration
                        }
                        op2 = op2.Next;
                    }
                    op = op.Next;
                }
                while (op != outrec.Pts);
            }
        }
        //------------------------------------------------------------------------------

        public static double Area(Path poly)
        {
            int cnt = (int)poly.Count;
            if (cnt < 3) return 0;
            double a = 0;
            for (int i = 0, j = cnt - 1; i < cnt; ++i)
            {
                a += ((double)poly[j].X + poly[i].X) * ((double)poly[j].Y - poly[i].Y);
                j = i;
            }
            return -a * 0.5;
        }
        //------------------------------------------------------------------------------

        internal double Area(OutRec outRec)
        {
            return Area(outRec.Pts);
        }
        //------------------------------------------------------------------------------

        internal double Area(OutPt op)
        {
            OutPt opFirst = op;
            if (op == null) return 0;
            double a = 0;
            do
            {
                a = a + (double)(op.Prev.Pt.X + op.Pt.X) * (double)(op.Prev.Pt.Y - op.Pt.Y);
                op = op.Next;
            } while (op != opFirst);
            return a * 0.5;
        }

        //------------------------------------------------------------------------------
        // SimplifyPolygon functions ...
        // Convert self-intersecting polygons into simple polygons
        //------------------------------------------------------------------------------

        public static Paths SimplifyPolygon(Path poly,
              PolyFillType fillType = PolyFillType.pftEvenOdd)
        {
            Paths result = new Paths();
            Clipper c = new Clipper();
            c.StrictlySimple = true;
            c.AddPath(poly, PolyType.ptSubject, true);
            c.Execute(ClipType.ctUnion, result, fillType, fillType);
            return result;
        }
        //------------------------------------------------------------------------------

        public static Paths SimplifyPolygons(Paths polys,
            PolyFillType fillType = PolyFillType.pftEvenOdd)
        {
            Paths result = new Paths();
            Clipper c = new Clipper();
            c.StrictlySimple = true;
            c.AddPaths(polys, PolyType.ptSubject, true);
            c.Execute(ClipType.ctUnion, result, fillType, fillType);
            return result;
        }
        //------------------------------------------------------------------------------

        private static double DistanceSqrd(IntPoint pt1, IntPoint pt2)
        {
            double dx = ((double)pt1.X - pt2.X);
            double dy = ((double)pt1.Y - pt2.Y);
            return (dx * dx + dy * dy);
        }
        //------------------------------------------------------------------------------

        private static double DistanceFromLineSqrd(IntPoint pt, IntPoint ln1, IntPoint ln2)
        {
            //The equation of a line in general form (Ax + By + C = 0)
            //given 2 points (x¹,y¹) & (x²,y²) is ...
            //(y¹ - y²)x + (x² - x¹)y + (y² - y¹)x¹ - (x² - x¹)y¹ = 0
            //A = (y¹ - y²); B = (x² - x¹); C = (y² - y¹)x¹ - (x² - x¹)y¹
            //perpendicular distance of point (x³,y³) = (Ax³ + By³ + C)/Sqrt(A² + B²)
            //see http://en.wikipedia.org/wiki/Perpendicular_distance
            double A = ln1.Y - ln2.Y;
            double B = ln2.X - ln1.X;
            double C = A * ln1.X + B * ln1.Y;
            C = A * pt.X + B * pt.Y - C;
            return (C * C) / (A * A + B * B);
        }
        //---------------------------------------------------------------------------

        private static bool SlopesNearCollinear(IntPoint pt1,
            IntPoint pt2, IntPoint pt3, double distSqrd)
        {
            //this function is more accurate when the point that's GEOMETRICALLY 
            //between the other 2 points is the one that's tested for distance.  
            //nb: with 'spikes', either pt1 or pt3 is geometrically between the other pts                    
            if (Math.Abs(pt1.X - pt2.X) > Math.Abs(pt1.Y - pt2.Y))
            {
                if ((pt1.X > pt2.X) == (pt1.X < pt3.X))
                    return DistanceFromLineSqrd(pt1, pt2, pt3) < distSqrd;
                else if ((pt2.X > pt1.X) == (pt2.X < pt3.X))
                    return DistanceFromLineSqrd(pt2, pt1, pt3) < distSqrd;
                else
                    return DistanceFromLineSqrd(pt3, pt1, pt2) < distSqrd;
            }
            else
            {
                if ((pt1.Y > pt2.Y) == (pt1.Y < pt3.Y))
                    return DistanceFromLineSqrd(pt1, pt2, pt3) < distSqrd;
                else if ((pt2.Y > pt1.Y) == (pt2.Y < pt3.Y))
                    return DistanceFromLineSqrd(pt2, pt1, pt3) < distSqrd;
                else
                    return DistanceFromLineSqrd(pt3, pt1, pt2) < distSqrd;
            }
        }
        //------------------------------------------------------------------------------

        private static bool PointsAreClose(IntPoint pt1, IntPoint pt2, double distSqrd)
        {
            double dx = (double)pt1.X - pt2.X;
            double dy = (double)pt1.Y - pt2.Y;
            return ((dx * dx) + (dy * dy) <= distSqrd);
        }
        //------------------------------------------------------------------------------

        private static OutPt ExcludeOp(OutPt op)
        {
            OutPt result = op.Prev;
            result.Next = op.Next;
            op.Next.Prev = result;
            result.Idx = 0;
            return result;
        }
        //------------------------------------------------------------------------------

        public static Path CleanPolygon(Path path, double distance = 1.415)
        {
            //distance = proximity in units/pixels below which vertices will be stripped. 
            //Default ~= sqrt(2) so when adjacent vertices or semi-adjacent vertices have 
            //both x & y coords within 1 unit, then the second vertex will be stripped.

            int cnt = path.Count;

            if (cnt == 0) return new Path();

            OutPt[] outPts = new OutPt[cnt];
            for (int i = 0; i < cnt; ++i) outPts[i] = new OutPt();

            for (int i = 0; i < cnt; ++i)
            {
                outPts[i].Pt = path[i];
                outPts[i].Next = outPts[(i + 1) % cnt];
                outPts[i].Next.Prev = outPts[i];
                outPts[i].Idx = 0;
            }

            double distSqrd = distance * distance;
            OutPt op = outPts[0];
            while (op.Idx == 0 && op.Next != op.Prev)
            {
                if (PointsAreClose(op.Pt, op.Prev.Pt, distSqrd))
                {
                    op = ExcludeOp(op);
                    cnt--;
                }
                else if (PointsAreClose(op.Prev.Pt, op.Next.Pt, distSqrd))
                {
                    ExcludeOp(op.Next);
                    op = ExcludeOp(op);
                    cnt -= 2;
                }
                else if (SlopesNearCollinear(op.Prev.Pt, op.Pt, op.Next.Pt, distSqrd))
                {
                    op = ExcludeOp(op);
                    cnt--;
                }
                else
                {
                    op.Idx = 1;
                    op = op.Next;
                }
            }

            if (cnt < 3) cnt = 0;
            Path result = new Path(cnt);
            for (int i = 0; i < cnt; ++i)
            {
                result.Add(op.Pt);
                op = op.Next;
            }
            outPts = null;
            return result;
        }
        //------------------------------------------------------------------------------

        public static Paths CleanPolygons(Paths polys,
            double distance = 1.415)
        {
            Paths result = new Paths(polys.Count);
            for (int i = 0; i < polys.Count; i++)
                result.Add(CleanPolygon(polys[i], distance));
            return result;
        }
        //------------------------------------------------------------------------------

        internal static Paths Minkowski(Path pattern, Path path, bool IsSum, bool IsClosed)
        {
            int delta = (IsClosed ? 1 : 0);
            int polyCnt = pattern.Count;
            int pathCnt = path.Count;
            Paths result = new Paths(pathCnt);
            if (IsSum)
                for (int i = 0; i < pathCnt; i++)
                {
                    Path p = new Path(polyCnt);
                    foreach (IntPoint ip in pattern)
                        p.Add(new IntPoint(path[i].X + ip.X, path[i].Y + ip.Y));
                    result.Add(p);
                }
            else
                for (int i = 0; i < pathCnt; i++)
                {
                    Path p = new Path(polyCnt);
                    foreach (IntPoint ip in pattern)
                        p.Add(new IntPoint(path[i].X - ip.X, path[i].Y - ip.Y));
                    result.Add(p);
                }

            Paths quads = new Paths((pathCnt + delta) * (polyCnt + 1));
            for (int i = 0; i < pathCnt - 1 + delta; i++)
                for (int j = 0; j < polyCnt; j++)
                {
                    Path quad = new Path(4);
                    quad.Add(result[i % pathCnt][j % polyCnt]);
                    quad.Add(result[(i + 1) % pathCnt][j % polyCnt]);
                    quad.Add(result[(i + 1) % pathCnt][(j + 1) % polyCnt]);
                    quad.Add(result[i % pathCnt][(j + 1) % polyCnt]);
                    if (!Orientation(quad)) quad.Reverse();
                    quads.Add(quad);
                }
            return quads;
        }
        //------------------------------------------------------------------------------

        public static Paths MinkowskiSum(Path pattern, Path path, bool pathIsClosed)
        {
            Paths paths = Minkowski(pattern, path, true, pathIsClosed);
            Clipper c = new Clipper();
            c.AddPaths(paths, PolyType.ptSubject, true);
            c.Execute(ClipType.ctUnion, paths, PolyFillType.pftNonZero, PolyFillType.pftNonZero);
            return paths;
        }
        //------------------------------------------------------------------------------

        private static Path TranslatePath(Path path, IntPoint delta)
        {
            Path outPath = new Path(path.Count);
            for (int i = 0; i < path.Count; i++)
                outPath.Add(new IntPoint(path[i].X + delta.X, path[i].Y + delta.Y));
            return outPath;
        }
        //------------------------------------------------------------------------------

        public static Paths MinkowskiSum(Path pattern, Paths paths, bool pathIsClosed)
        {
            Paths solution = new Paths();
            Clipper c = new Clipper();
            for (int i = 0; i < paths.Count; ++i)
            {
                Paths tmp = Minkowski(pattern, paths[i], true, pathIsClosed);
                c.AddPaths(tmp, PolyType.ptSubject, true);
                if (pathIsClosed)
                {
                    Path path = TranslatePath(paths[i], pattern[0]);
                    c.AddPath(path, PolyType.ptClip, true);
                }
            }
            c.Execute(ClipType.ctUnion, solution,
              PolyFillType.pftNonZero, PolyFillType.pftNonZero);
            return solution;
        }
        //------------------------------------------------------------------------------

        public static Paths MinkowskiDiff(Path poly1, Path poly2)
        {
            Paths paths = Minkowski(poly1, poly2, false, true);
            Clipper c = new Clipper();
            c.AddPaths(paths, PolyType.ptSubject, true);
            c.Execute(ClipType.ctUnion, paths, PolyFillType.pftNonZero, PolyFillType.pftNonZero);
            return paths;
        }
        //------------------------------------------------------------------------------

        internal enum NodeType { ntAny, ntOpen, ntClosed };

        public static Paths PolyTreeToPaths(PolyTree polytree)
        {

            Paths result = new Paths();
            result.Capacity = polytree.Total;
            AddPolyNodeToPaths(polytree, NodeType.ntAny, result);
            return result;
        }
        //------------------------------------------------------------------------------

        internal static void AddPolyNodeToPaths(PolyNode polynode, NodeType nt, Paths paths)
        {
            bool match = true;
            switch (nt)
            {
                case NodeType.ntOpen: return;
                case NodeType.ntClosed: match = !polynode.IsOpen; break;
                default: break;
            }

            if (polynode.m_polygon.Count > 0 && match)
                paths.Add(polynode.m_polygon);
            foreach (PolyNode pn in polynode.Childs)
                AddPolyNodeToPaths(pn, nt, paths);
        }
        //------------------------------------------------------------------------------

        public static Paths OpenPathsFromPolyTree(PolyTree polytree)
        {
            Paths result = new Paths();
            result.Capacity = polytree.ChildCount;
            for (int i = 0; i < polytree.ChildCount; i++)
                if (polytree.Childs[i].IsOpen)
                    result.Add(polytree.Childs[i].m_polygon);
            return result;
        }
        //------------------------------------------------------------------------------

        public static Paths ClosedPathsFromPolyTree(PolyTree polytree)
        {
            Paths result = new Paths();
            result.Capacity = polytree.Total;
            AddPolyNodeToPaths(polytree, NodeType.ntClosed, result);
            return result;
        }
        //------------------------------------------------------------------------------

    } //end Clipper

    public class ClipperOffset
    {
        private Paths m_destPolys;
        private Path m_srcPoly;
        private Path m_destPoly;
        private List<DoublePoint> m_normals = new List<DoublePoint>();
        private double m_delta, m_sinA, m_sin, m_cos;
        private double m_miterLim, m_StepsPerRad;

        private IntPoint m_lowest;
        private PolyNode m_polyNodes = new PolyNode();

        public double ArcTolerance { get; set; }
        public double MiterLimit { get; set; }

        private const double two_pi = Math.PI * 2;
        private const double def_arc_tolerance = 0.25;

        public ClipperOffset(
          double miterLimit = 2.0, double arcTolerance = def_arc_tolerance)
        {
            MiterLimit = miterLimit;
            ArcTolerance = arcTolerance;
            m_lowest.X = -1;
        }
        //------------------------------------------------------------------------------

        public void Clear()
        {
            m_polyNodes.Childs.Clear();
            m_lowest.X = -1;
        }
        //------------------------------------------------------------------------------

        internal static cInt Round(double value)
        {
            return value < 0 ? (cInt)(value - 0.5) : (cInt)(value + 0.5);
        }
        //------------------------------------------------------------------------------

        public void AddPath(Path path, JoinType joinType, EndType endType)
        {
            int highI = path.Count - 1;
            if (highI < 0) return;
            PolyNode newNode = new PolyNode();
            newNode.m_jointype = joinType;
            newNode.m_endtype = endType;

            //strip duplicate points from path and also get index to the lowest point ...
            if (endType == EndType.etClosedLine || endType == EndType.etClosedPolygon)
                while (highI > 0 && path[0] == path[highI]) highI--;
            newNode.m_polygon.Capacity = highI + 1;
            newNode.m_polygon.Add(path[0]);
            int j = 0, k = 0;
            for (int i = 1; i <= highI; i++)
                if (newNode.m_polygon[j] != path[i])
                {
                    j++;
                    newNode.m_polygon.Add(path[i]);
                    if (path[i].Y > newNode.m_polygon[k].Y ||
                      (path[i].Y == newNode.m_polygon[k].Y &&
                      path[i].X < newNode.m_polygon[k].X)) k = j;
                }
            if (endType == EndType.etClosedPolygon && j < 2) return;

            m_polyNodes.AddChild(newNode);

            //if this path's lowest pt is lower than all the others then update m_lowest
            if (endType != EndType.etClosedPolygon) return;
            if (m_lowest.X < 0)
                m_lowest = new IntPoint(m_polyNodes.ChildCount - 1, k);
            else
            {
                IntPoint ip = m_polyNodes.Childs[(int)m_lowest.X].m_polygon[(int)m_lowest.Y];
                if (newNode.m_polygon[k].Y > ip.Y ||
                  (newNode.m_polygon[k].Y == ip.Y &&
                  newNode.m_polygon[k].X < ip.X))
                    m_lowest = new IntPoint(m_polyNodes.ChildCount - 1, k);
            }
        }
        //------------------------------------------------------------------------------

        public void AddPaths(Paths paths, JoinType joinType, EndType endType)
        {
            foreach (Path p in paths)
                AddPath(p, joinType, endType);
        }
        //------------------------------------------------------------------------------

        private void FixOrientations()
        {
            //fixup orientations of all closed paths if the orientation of the
            //closed path with the lowermost vertex is wrong ...
            if (m_lowest.X >= 0 &&
              !Clipper.Orientation(m_polyNodes.Childs[(int)m_lowest.X].m_polygon))
            {
                for (int i = 0; i < m_polyNodes.ChildCount; i++)
                {
                    PolyNode node = m_polyNodes.Childs[i];
                    if (node.m_endtype == EndType.etClosedPolygon ||
                      (node.m_endtype == EndType.etClosedLine &&
                      Clipper.Orientation(node.m_polygon)))
                        node.m_polygon.Reverse();
                }
            }
            else
            {
                for (int i = 0; i < m_polyNodes.ChildCount; i++)
                {
                    PolyNode node = m_polyNodes.Childs[i];
                    if (node.m_endtype == EndType.etClosedLine &&
                      !Clipper.Orientation(node.m_polygon))
                        node.m_polygon.Reverse();
                }
            }
        }
        //------------------------------------------------------------------------------

        internal static DoublePoint GetUnitNormal(IntPoint pt1, IntPoint pt2)
        {
            double dx = (pt2.X - pt1.X);
            double dy = (pt2.Y - pt1.Y);
            if ((dx == 0) && (dy == 0)) return new DoublePoint();

            double f = 1 * 1.0 / Math.Sqrt(dx * dx + dy * dy);
            dx *= f;
            dy *= f;

            return new DoublePoint(dy, -dx);
        }
        //------------------------------------------------------------------------------

        private void DoOffset(double delta)
        {
            m_destPolys = new Paths();
            m_delta = delta;

            //if Zero offset, just copy any CLOSED polygons to m_p and return ...
            if (ClipperBase.near_zero(delta))
            {
                m_destPolys.Capacity = m_polyNodes.ChildCount;
                for (int i = 0; i < m_polyNodes.ChildCount; i++)
                {
                    PolyNode node = m_polyNodes.Childs[i];
                    if (node.m_endtype == EndType.etClosedPolygon)
                        m_destPolys.Add(node.m_polygon);
                }
                return;
            }

            //see offset_triginometry3.svg in the documentation folder ...
            if (MiterLimit > 2) m_miterLim = 2 / (MiterLimit * MiterLimit);
            else m_miterLim = 0.5;

            double y;
            if (ArcTolerance <= 0.0)
                y = def_arc_tolerance;
            else if (ArcTolerance > Math.Abs(delta) * def_arc_tolerance)
                y = Math.Abs(delta) * def_arc_tolerance;
            else
                y = ArcTolerance;
            //see offset_triginometry2.svg in the documentation folder ...
            double steps = Math.PI / Math.Acos(1 - y / Math.Abs(delta));
            m_sin = Math.Sin(two_pi / steps);
            m_cos = Math.Cos(two_pi / steps);
            m_StepsPerRad = steps / two_pi;
            if (delta < 0.0) m_sin = -m_sin;

            m_destPolys.Capacity = m_polyNodes.ChildCount * 2;
            for (int i = 0; i < m_polyNodes.ChildCount; i++)
            {
                PolyNode node = m_polyNodes.Childs[i];
                m_srcPoly = node.m_polygon;

                int len = m_srcPoly.Count;

                if (len == 0 || (delta <= 0 && (len < 3 ||
                  node.m_endtype != EndType.etClosedPolygon)))
                    continue;

                m_destPoly = new Path();

                if (len == 1)
                {
                    if (node.m_jointype == JoinType.jtRound)
                    {
                        double X = 1.0, Y = 0.0;
                        for (int j = 1; j <= steps; j++)
                        {
                            m_destPoly.Add(new IntPoint(
                              Round(m_srcPoly[0].X + X * delta),
                              Round(m_srcPoly[0].Y + Y * delta)));
                            double X2 = X;
                            X = X * m_cos - m_sin * Y;
                            Y = X2 * m_sin + Y * m_cos;
                        }
                    }
                    else
                    {
                        double X = -1.0, Y = -1.0;
                        for (int j = 0; j < 4; ++j)
                        {
                            m_destPoly.Add(new IntPoint(
                              Round(m_srcPoly[0].X + X * delta),
                              Round(m_srcPoly[0].Y + Y * delta)));
                            if (X < 0) X = 1;
                            else if (Y < 0) Y = 1;
                            else X = -1;
                        }
                    }
                    m_destPolys.Add(m_destPoly);
                    continue;
                }

                //build m_normals ...
                m_normals.Clear();
                m_normals.Capacity = len;
                for (int j = 0; j < len - 1; j++)
                    m_normals.Add(GetUnitNormal(m_srcPoly[j], m_srcPoly[j + 1]));
                if (node.m_endtype == EndType.etClosedLine ||
                  node.m_endtype == EndType.etClosedPolygon)
                    m_normals.Add(GetUnitNormal(m_srcPoly[len - 1], m_srcPoly[0]));
                else
                    m_normals.Add(new DoublePoint(m_normals[len - 2]));

                if (node.m_endtype == EndType.etClosedPolygon)
                {
                    int k = len - 1;
                    for (int j = 0; j < len; j++)
                        OffsetPoint(j, ref k, node.m_jointype);
                    m_destPolys.Add(m_destPoly);
                }
                else if (node.m_endtype == EndType.etClosedLine)
                {
                    int k = len - 1;
                    for (int j = 0; j < len; j++)
                        OffsetPoint(j, ref k, node.m_jointype);
                    m_destPolys.Add(m_destPoly);
                    m_destPoly = new Path();
                    //re-build m_normals ...
                    DoublePoint n = m_normals[len - 1];
                    for (int j = len - 1; j > 0; j--)
                        m_normals[j] = new DoublePoint(-m_normals[j - 1].X, -m_normals[j - 1].Y);
                    m_normals[0] = new DoublePoint(-n.X, -n.Y);
                    k = 0;
                    for (int j = len - 1; j >= 0; j--)
                        OffsetPoint(j, ref k, node.m_jointype);
                    m_destPolys.Add(m_destPoly);
                }
                else
                {
                    int k = 0;
                    for (int j = 1; j < len - 1; ++j)
                        OffsetPoint(j, ref k, node.m_jointype);

                    IntPoint pt1;
                    if (node.m_endtype == EndType.etOpenButt)
                    {
                        int j = len - 1;
                        pt1 = new IntPoint((cInt)Round(m_srcPoly[j].X + m_normals[j].X *
                          delta), (cInt)Round(m_srcPoly[j].Y + m_normals[j].Y * delta));
                        m_destPoly.Add(pt1);
                        pt1 = new IntPoint((cInt)Round(m_srcPoly[j].X - m_normals[j].X *
                          delta), (cInt)Round(m_srcPoly[j].Y - m_normals[j].Y * delta));
                        m_destPoly.Add(pt1);
                    }
                    else
                    {
                        int j = len - 1;
                        k = len - 2;
                        m_sinA = 0;
                        m_normals[j] = new DoublePoint(-m_normals[j].X, -m_normals[j].Y);
                        if (node.m_endtype == EndType.etOpenSquare)
                            DoSquare(j, k);
                        else
                            DoRound(j, k);
                    }

                    //re-build m_normals ...
                    for (int j = len - 1; j > 0; j--)
                        m_normals[j] = new DoublePoint(-m_normals[j - 1].X, -m_normals[j - 1].Y);

                    m_normals[0] = new DoublePoint(-m_normals[1].X, -m_normals[1].Y);

                    k = len - 1;
                    for (int j = k - 1; j > 0; --j)
                        OffsetPoint(j, ref k, node.m_jointype);

                    if (node.m_endtype == EndType.etOpenButt)
                    {
                        pt1 = new IntPoint((cInt)Round(m_srcPoly[0].X - m_normals[0].X * delta),
                          (cInt)Round(m_srcPoly[0].Y - m_normals[0].Y * delta));
                        m_destPoly.Add(pt1);
                        pt1 = new IntPoint((cInt)Round(m_srcPoly[0].X + m_normals[0].X * delta),
                          (cInt)Round(m_srcPoly[0].Y + m_normals[0].Y * delta));
                        m_destPoly.Add(pt1);
                    }
                    else
                    {
                        k = 1;
                        m_sinA = 0;
                        if (node.m_endtype == EndType.etOpenSquare)
                            DoSquare(0, 1);
                        else
                            DoRound(0, 1);
                    }
                    m_destPolys.Add(m_destPoly);
                }
            }
        }
        //------------------------------------------------------------------------------

        public void Execute(ref Paths solution, double delta)
        {
            solution.Clear();
            FixOrientations();
            DoOffset(delta);
            //now clean up 'corners' ...
            Clipper clpr = new Clipper();
            clpr.AddPaths(m_destPolys, PolyType.ptSubject, true);
            if (delta > 0)
            {
                clpr.Execute(ClipType.ctUnion, solution,
                  PolyFillType.pftPositive, PolyFillType.pftPositive);
            }
            else
            {
                IntRect r = Clipper.GetBounds(m_destPolys);
                Path outer = new Path(4);

                outer.Add(new IntPoint(r.left - 10, r.bottom + 10));
                outer.Add(new IntPoint(r.right + 10, r.bottom + 10));
                outer.Add(new IntPoint(r.right + 10, r.top - 10));
                outer.Add(new IntPoint(r.left - 10, r.top - 10));

                clpr.AddPath(outer, PolyType.ptSubject, true);
                clpr.ReverseSolution = true;
                clpr.Execute(ClipType.ctUnion, solution, PolyFillType.pftNegative, PolyFillType.pftNegative);
                if (solution.Count > 0) solution.RemoveAt(0);
            }
        }
        //------------------------------------------------------------------------------

        public void Execute(ref PolyTree solution, double delta)
        {
            solution.Clear();
            FixOrientations();
            DoOffset(delta);

            //now clean up 'corners' ...
            Clipper clpr = new Clipper();
            clpr.AddPaths(m_destPolys, PolyType.ptSubject, true);
            if (delta > 0)
            {
                clpr.Execute(ClipType.ctUnion, solution,
                  PolyFillType.pftPositive, PolyFillType.pftPositive);
            }
            else
            {
                IntRect r = Clipper.GetBounds(m_destPolys);
                Path outer = new Path(4);

                outer.Add(new IntPoint(r.left - 10, r.bottom + 10));
                outer.Add(new IntPoint(r.right + 10, r.bottom + 10));
                outer.Add(new IntPoint(r.right + 10, r.top - 10));
                outer.Add(new IntPoint(r.left - 10, r.top - 10));

                clpr.AddPath(outer, PolyType.ptSubject, true);
                clpr.ReverseSolution = true;
                clpr.Execute(ClipType.ctUnion, solution, PolyFillType.pftNegative, PolyFillType.pftNegative);
                //remove the outer PolyNode rectangle ...
                if (solution.ChildCount == 1 && solution.Childs[0].ChildCount > 0)
                {
                    PolyNode outerNode = solution.Childs[0];
                    solution.Childs.Capacity = outerNode.ChildCount;
                    solution.Childs[0] = outerNode.Childs[0];
                    solution.Childs[0].m_Parent = solution;
                    for (int i = 1; i < outerNode.ChildCount; i++)
                        solution.AddChild(outerNode.Childs[i]);
                }
                else
                    solution.Clear();
            }
        }
        //------------------------------------------------------------------------------

        void OffsetPoint(int j, ref int k, JoinType jointype)
        {
            //cross product ...
            m_sinA = (m_normals[k].X * m_normals[j].Y - m_normals[j].X * m_normals[k].Y);

            if (Math.Abs(m_sinA * m_delta) < 1.0)
            {
                //dot product ...
                double cosA = (m_normals[k].X * m_normals[j].X + m_normals[j].Y * m_normals[k].Y);
                if (cosA > 0) // angle ==> 0 degrees
                {
                    m_destPoly.Add(new IntPoint(Round(m_srcPoly[j].X + m_normals[k].X * m_delta),
                      Round(m_srcPoly[j].Y + m_normals[k].Y * m_delta)));
                    return;
                }
                //else angle ==> 180 degrees   
            }
            else if (m_sinA > 1.0) m_sinA = 1.0;
            else if (m_sinA < -1.0) m_sinA = -1.0;

            if (m_sinA * m_delta < 0)
            {
                m_destPoly.Add(new IntPoint(Round(m_srcPoly[j].X + m_normals[k].X * m_delta),
                  Round(m_srcPoly[j].Y + m_normals[k].Y * m_delta)));
                m_destPoly.Add(m_srcPoly[j]);
                m_destPoly.Add(new IntPoint(Round(m_srcPoly[j].X + m_normals[j].X * m_delta),
                  Round(m_srcPoly[j].Y + m_normals[j].Y * m_delta)));
            }
            else
                switch (jointype)
                {
                    case JoinType.jtMiter:
                        {
                            double r = 1 + (m_normals[j].X * m_normals[k].X +
                              m_normals[j].Y * m_normals[k].Y);
                            if (r >= m_miterLim) DoMiter(j, k, r); else DoSquare(j, k);
                            break;
                        }
                    case JoinType.jtSquare: DoSquare(j, k); break;
                    case JoinType.jtRound: DoRound(j, k); break;
                }
            k = j;
        }
        //------------------------------------------------------------------------------

        internal void DoSquare(int j, int k)
        {
            double dx = Math.Tan(Math.Atan2(m_sinA,
                m_normals[k].X * m_normals[j].X + m_normals[k].Y * m_normals[j].Y) / 4);
            m_destPoly.Add(new IntPoint(
                Round(m_srcPoly[j].X + m_delta * (m_normals[k].X - m_normals[k].Y * dx)),
                Round(m_srcPoly[j].Y + m_delta * (m_normals[k].Y + m_normals[k].X * dx))));
            m_destPoly.Add(new IntPoint(
                Round(m_srcPoly[j].X + m_delta * (m_normals[j].X + m_normals[j].Y * dx)),
                Round(m_srcPoly[j].Y + m_delta * (m_normals[j].Y - m_normals[j].X * dx))));
        }
        //------------------------------------------------------------------------------

        internal void DoMiter(int j, int k, double r)
        {
            double q = m_delta / r;
            m_destPoly.Add(new IntPoint(Round(m_srcPoly[j].X + (m_normals[k].X + m_normals[j].X) * q),
                Round(m_srcPoly[j].Y + (m_normals[k].Y + m_normals[j].Y) * q)));
        }
        //------------------------------------------------------------------------------

        internal void DoRound(int j, int k)
        {
            double a = Math.Atan2(m_sinA,
            m_normals[k].X * m_normals[j].X + m_normals[k].Y * m_normals[j].Y);
            int steps = Math.Max((int)Round(m_StepsPerRad * Math.Abs(a)), 1);

            double X = m_normals[k].X, Y = m_normals[k].Y, X2;
            for (int i = 0; i < steps; ++i)
            {
                m_destPoly.Add(new IntPoint(
                    Round(m_srcPoly[j].X + X * m_delta),
                    Round(m_srcPoly[j].Y + Y * m_delta)));
                X2 = X;
                X = X * m_cos - m_sin * Y;
                Y = X2 * m_sin + Y * m_cos;
            }
            m_destPoly.Add(new IntPoint(
            Round(m_srcPoly[j].X + m_normals[j].X * m_delta),
            Round(m_srcPoly[j].Y + m_normals[j].Y * m_delta)));
        }
        //------------------------------------------------------------------------------
    }

    class ClipperException : Exception
    {
        public ClipperException(string description) : base(description) { }
    }
    //------------------------------------------------------------------------------

} //end ClipperLib namespace

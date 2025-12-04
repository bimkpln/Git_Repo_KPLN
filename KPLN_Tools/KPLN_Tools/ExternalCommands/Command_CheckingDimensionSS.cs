using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    internal class Command_CheckingDimensionSS : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication _uiapp = commandData.Application;
            UIDocument _uidoc = _uiapp.ActiveUIDocument;
            Document _doc = _uidoc.Document;

            TaskDialog td = new TaskDialog("Проверка габаритов электрооборудования и элементов узлов");
            td.MainInstruction = "Что будем проверять?";
            td.MainContent = "Выберите одно из действий ниже.";
            td.CommonButtons = TaskDialogCommonButtons.Close;
            td.DefaultButton = TaskDialogResult.Close;

            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Проверка параметров");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Проверка габаритов");

            TaskDialogResult result = td.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                CheckGroupingVsPanelName(_uiapp, _doc);
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                CheckDimensionsForGroups(_uiapp, _doc);
            }

            return Result.Succeeded;
        }










        private void CheckGroupingVsPanelName(UIApplication uiapp, Document doc)
        {
            // Элементы узлов
            string[] nodeFamilyNames =
            {
                "076_КШ_Шкаф_Универсальный_(ЭлУзл)",
                "076_КШ_Шкаф_Корпус ЩМП У2 IP54_(ЭлУзл)"
            };

            var nodeElements = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    Family family = fi.Symbol?.Family;
                    return family != null && nodeFamilyNames.Contains(family.Name);
                })
                .Cast<Element>()
                .ToList();

            if (!nodeElements.Any())
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Не найдено ни одного элемента узла указанных семейств.");
                return;
            }

            const string groupParamName = "КП_О_Группирование";
            const string panelNameParam = "Имя панели";

            // Значения "КП_О_Группирование"
            var groupingValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Element e in nodeElements)
            {
                string val = GetStringParam(e, groupParamName);
                if (!string.IsNullOrWhiteSpace(val))
                    groupingValues.Add(val.Trim());
            }

            if (!groupingValues.Any())
            {
                TaskDialog.Show("KPLN. Результаты проверки", $"У элементов узлов не найдено ни одного непустого значения параметра \"{groupParamName}\".");
                return;
            }

            // Электрооборудование 
            var cabinetInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name == "851_Щит_Универсальный_(ЭлОб)")
                .ToList();

            if (!cabinetInstances.Any())
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Не найдено ни одного элемента категории OST_ElectricalEquipment.");
                return;
            }

            // Множество всех "Имя панели"
            var panelNamesSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyInstance fi in cabinetInstances)
            {
                string panelName = GetStringParam(fi, panelNameParam);
                if (!string.IsNullOrWhiteSpace(panelName))
                    panelNamesSet.Add(panelName.Trim());
            }

            // Электрооборудование без соответствия (Имя панели ∉ КП_О_Группирование)
            var cabinetMismatches = new List<CabinetMismatchItem>();
            foreach (FamilyInstance fi in cabinetInstances)
            {
                string panelName = GetStringParam(fi, panelNameParam)?.Trim();

                if (string.IsNullOrWhiteSpace(panelName) || !groupingValues.Contains(panelName))
                {
                    cabinetMismatches.Add(new CabinetMismatchItem
                    {
                        ElementId = fi.Id,
                        FamilyName = fi.Symbol?.Family?.Name ?? "<без имени семейства>",
                        TypeName = fi.Symbol?.Name ?? "<без типа>",
                        PanelName = panelName ?? "<пусто>"
                    });
                }
            }

            // Элементы узлов без соответствия (КП_О_Группирование ∉ Имя панели)
            var nodeMismatches = new List<NodeMismatchItem>();
            foreach (Element e in nodeElements)
            {
                string gv = GetStringParam(e, groupParamName)?.Trim();
                if (string.IsNullOrWhiteSpace(gv))
                    continue;

                if (!panelNamesSet.Contains(gv))
                {
                    var fi = e as FamilyInstance;
                    nodeMismatches.Add(new NodeMismatchItem
                    {
                        ElementId = e.Id,
                        FamilyName = fi?.Symbol?.Family?.Name ?? "<без имени семейства>",
                        TypeName = fi?.Symbol?.Name ?? "<без типа>",
                        GroupingValue = gv
                    });
                }
            }

            if (!cabinetMismatches.Any() && !nodeMismatches.Any())
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Несоответствий между \"КП_О_Группирование\" и \"Имя панели\" не найдено.");
                return;
            }

            // Окно со списками
            var window = new Forms.CabinetMismatchWindow(uiapp, cabinetMismatches, nodeMismatches);
            window.Show();
        }

        private static string GetStringParam(Element e, string paramName)
        {
            Parameter p = e.LookupParameter(paramName);

            if (p == null)
            {
                FamilyInstance fi = e as FamilyInstance;
                if (fi != null && fi.Symbol != null)
                {
                    p = fi.Symbol.LookupParameter(paramName);
                }
            }

            return p != null ? p.AsString() : null;
        }











        private void CheckDimensionsForGroups(UIApplication uiapp, Document doc)
        {
            // Имена семейств
            const string famNode1 = "076_КШ_Шкаф_Универсальный_(ЭлУзл)";
            const string famNode2 = "076_КШ_Шкаф_Корпус ЩМП У2 IP54_(ЭлУзл)";
            const string famCabinet = "851_Щит_Универсальный_(ЭлОб)";

            const string groupParamName = "КП_О_Группирование";
            const string panelNameParam = "Имя панели";

            // Собираем все экземпляры нужных семейств
            IList<FamilyInstance> allFi = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            // Элементы узлов
            List<FamilyInstance> nodeInstances = new List<FamilyInstance>();
            // Электрооборудование
            List<FamilyInstance> cabinetInstances = new List<FamilyInstance>();

            foreach (FamilyInstance fi in allFi)
            {
                Family family = fi.Symbol != null ? fi.Symbol.Family : null;
                if (family == null) { continue; }

                string famName = family.Name;

                if (famName == famNode1 || famName == famNode2)
                {
                    nodeInstances.Add(fi);
                }
                else if (famName == famCabinet)
                {
                    cabinetInstances.Add(fi);
                }
            }

            if (nodeInstances.Count == 0)
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Не найдено ни одного элемента узлов указанных семейств.");
                return;
            }

            if (cabinetInstances.Count == 0)
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Не найдено ни одного элемента семейства 851_Щит_Универсальный_(ЭлОб).");
                return;
            }

            // Группируем узлы по "КП_О_Группирование"
            Dictionary<string, List<FamilyInstance>> nodesByKey =
                new Dictionary<string, List<FamilyInstance>>(StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance fi in nodeInstances)
            {
                string key = GetStringParam(fi, groupParamName);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                key = key.Trim();

                List<FamilyInstance> list;
                if (!nodesByKey.TryGetValue(key, out list))
                {
                    list = new List<FamilyInstance>();
                    nodesByKey.Add(key, list);
                }
                list.Add(fi);
            }

            // Группируем шкафы по "Имя панели"
            Dictionary<string, List<FamilyInstance>> cabinetsByKey =
                new Dictionary<string, List<FamilyInstance>>(StringComparer.OrdinalIgnoreCase);

            foreach (FamilyInstance fi in cabinetInstances)
            {
                string key = GetStringParam(fi, panelNameParam);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                key = key.Trim();

                List<FamilyInstance> list;
                if (!cabinetsByKey.TryGetValue(key, out list))
                {
                    list = new List<FamilyInstance>();
                    cabinetsByKey.Add(key, list);
                }
                list.Add(fi);
            }

            // Берём только те ключи, где есть и узлы, и шкафы (КП_О_Группирование = Имя панели)
            List<string> commonKeys = new List<string>();
            foreach (string k in nodesByKey.Keys)
            {
                if (cabinetsByKey.ContainsKey(k))
                {
                    commonKeys.Add(k);
                }
            }

            if (commonKeys.Count == 0)
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Не найдено ни одной группы, где \"КП_О_Группирование\" совпадает с \"Имя панели\".");
                return;
            }

            const double ftToM = 0.3048;
            const double tol = 1e-6;

            List<SizeMismatchItem> mismatches = new List<SizeMismatchItem>();

            // По каждой группе сравниваем габариты
            foreach (string key in commonKeys)
            {
                List<FamilyInstance> groupNodes = nodesByKey[key];
                List<FamilyInstance> groupCabs = cabinetsByKey[key];

                // Шкаф
                FamilyInstance refCab = groupCabs[0];

                double refH_ft, refW_ft, refD_ft;
                GetCabinetDimensions(refCab, out refH_ft, out refW_ft, out refD_ft);

                // Проверяем все шкафы группы
                foreach (FamilyInstance cab in groupCabs)
                {
                    double h_ft, w_ft, d_ft;
                    GetCabinetDimensions(cab, out h_ft, out w_ft, out d_ft);

                    bool diff =
                        IsDiff(h_ft, refH_ft, tol) ||
                        IsDiff(w_ft, refW_ft, tol) ||
                        IsDiff(d_ft, refD_ft, tol);

                    if (diff)
                    {
                        Family family = cab.Symbol != null ? cab.Symbol.Family : null;
                        string famName = family != null ? family.Name : "<без имени семейства>";
                        string typeName = cab.Symbol != null ? cab.Symbol.Name : "<без типа>";

                        mismatches.Add(new SizeMismatchItem
                        {
                            ElementId = cab.Id,
                            FamilyName = famName,
                            TypeName = typeName,
                            GroupKey = key,
                            IsCabinet = true,
                            HeightM = h_ft * ftToM,
                            WidthM = w_ft * ftToM,
                            DepthM = d_ft * ftToM,
                            RefHeightM = refH_ft * ftToM,
                            RefWidthM = refW_ft * ftToM,
                            RefDepthM = refD_ft * ftToM
                        });
                    }
                }

                // Проверяем все узлы группы
                foreach (FamilyInstance node in groupNodes)
                {
                    double h_ft, w_ft, d_ft;
                    GetNodeDimensions(node, out h_ft, out w_ft, out d_ft);

                    Family family = node.Symbol != null ? node.Symbol.Family : null;
                    string famName = family != null ? family.Name : string.Empty;

                    bool ignoreDepth =
                        string.Equals(famName, famNode1, StringComparison.OrdinalIgnoreCase); 

                    bool diff =
                        IsDiff(h_ft, refH_ft, tol) ||
                        IsDiff(w_ft, refW_ft, tol) ||
                        (!ignoreDepth && IsDiff(d_ft, refD_ft, tol));

                    if (diff)
                    {
                        string showFamName = famName != string.Empty ? famName : "<без имени семейства>";
                        string typeName = node.Symbol != null ? node.Symbol.Name : "<без типа>";

                        mismatches.Add(new SizeMismatchItem
                        {
                            ElementId = node.Id,
                            FamilyName = showFamName,
                            TypeName = typeName,
                            GroupKey = key,
                            IsCabinet = false,
                            HeightM = h_ft * ftToM,
                            WidthM = w_ft * ftToM,
                            DepthM = d_ft * ftToM,
                            RefHeightM = refH_ft * ftToM,
                            RefWidthM = refW_ft * ftToM,
                            RefDepthM = refD_ft * ftToM
                        });
                    }
                }
            }

            if (mismatches.Count == 0)
            {
                TaskDialog.Show("KPLN. Результаты проверки", "Во всех группах (КП_О_Группирование = Имя панели) габариты совпадают.");
                return;
            }

            // Окно с результатами
            KPLN_Tools.Forms.SizeMismatchWindow window =
                new KPLN_Tools.Forms.SizeMismatchWindow(uiapp, mismatches);
            window.Show();
        }

        private static double GetDoubleParam(Element e, string paramName)
        {
            Parameter p = e.LookupParameter(paramName);

            if (p == null)
            {
                FamilyInstance fi = e as FamilyInstance;
                if (fi != null && fi.Symbol != null)
                {
                    p = fi.Symbol.LookupParameter(paramName);
                }
            }

            if (p == null)
                return 0;

            return p.AsDouble();
        }

        private static bool IsDiff(double a, double b, double tol)
        {
            return Math.Abs(a - b) > tol;
        }

        /// <summary>
        /// Габариты шкафа 851_Щит_Универсальный_(ЭлОб)
        /// Корпус_Высота, Корпус_Ширина, Корпус_Глубина
        /// </summary>
        private static void GetCabinetDimensions(FamilyInstance fi,
            out double h_ft, out double w_ft, out double d_ft)
        {
            h_ft = GetDoubleParam(fi, "Корпус_Высота");
            w_ft = GetDoubleParam(fi, "Корпус_Ширина");
            d_ft = GetDoubleParam(fi, "Корпус_Глубина");
        }

        /// <summary>
        /// Габариты узлов:
        /// 1) 076_КШ_Шкаф_Универсальный_(ЭлУзл): Высота мм, Ширина шкафа (глубины нет)
        /// 2) 076_КШ_Шкаф_Корпус ЩМП У2 IP54_(ЭлУзл): Высота H, Глубина Y, Ширина W
        /// </summary>
        private static void GetNodeDimensions(FamilyInstance fi,
            out double h_ft, out double w_ft, out double d_ft)
        {
            h_ft = 0;
            w_ft = 0;
            d_ft = 0;

            Family family = fi.Symbol != null ? fi.Symbol.Family : null;
            string famName = family != null ? family.Name : string.Empty;

            if (famName == "076_КШ_Шкаф_Универсальный_(ЭлУзл)")
            {
                h_ft = GetDoubleParam(fi, "Высота мм");
                w_ft = GetDoubleParam(fi, "Ширина шкафа");
                d_ft = 0;
            }
            else if (famName == "076_КШ_Шкаф_Корпус ЩМП У2 IP54_(ЭлУзл)")
            {
                h_ft = GetDoubleParam(fi, "Высота H");
                d_ft = GetDoubleParam(fi, "Глубина Y");
                w_ft = GetDoubleParam(fi, "Ширина W");
            }
        }
    }


    public class CabinetMismatchItem
    {
        public ElementId ElementId { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string PanelName { get; set; }

        public string DisplayName =>
            $"[{ElementId.IntegerValue}] {FamilyName}\n{TypeName} ({PanelName})";
    }

    public class NodeMismatchItem
    {
        public ElementId ElementId { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string GroupingValue { get; set; }

        public string DisplayName =>
            $"[{ElementId.IntegerValue}] {FamilyName}\n{TypeName} ({GroupingValue})";
    }

    public class SizeMismatchItem
    {
        public ElementId ElementId { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }

        public string GroupKey { get; set; }

        public bool IsCabinet { get; set; }

        // Фактические габариты элемента в метрах
        public double HeightM { get; set; }
        public double WidthM { get; set; }
        public double DepthM { get; set; }

        // Габариты по группе в метрах
        public double RefHeightM { get; set; }
        public double RefWidthM { get; set; }
        public double RefDepthM { get; set; }

        public string DisplayName
        {
            get
            {
                string kind = IsCabinet ? "Шкаф" : "ЭлУзел";
                return
                    $"[{ElementId.IntegerValue}] {kind} {FamilyName}\n" +
                    $"{TypeName} (КП_О_Группирование: {GroupKey})\n" +
                    $"ЭлУзлов:   H={100*HeightM:0.###}  W={100*WidthM:0.###}  D={100*DepthM:0.###} м\n" +
                    $"Шкаф:       H={100*RefHeightM:0.###}  W={100*RefWidthM:0.###}  D={100*RefDepthM:0.###} м";
            }
        }
    }
}

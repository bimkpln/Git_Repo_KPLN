using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [Transaction(TransactionMode.Manual)]
    public class NonModelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                // Проверка проекта до открытия настроек и изменения модели.
                if (!Util.IsSetunProject(c))
                    return Result.Cancelled;

                var d = Util.Project(c);
                NonModelSettings s;


                // Настройка этапов обработки.
                using (var form = new NonModelForm(d))
                {
                    if (form.ShowDialog(new WindowOwner(c.Application)) != F.DialogResult.OK)
                        return Result.Cancelled;

                    s = form.Settings;
                }

                var report = new Report();
                report.Lines.Add("Немоделируемые элементы. Проект: " + d.Title);
                report.Lines.Add("Режим: " + (s.Preview ? "ПРОВЕРКА, без записи" : "ЗАПИСЬ"));

                // Подготовка групп шкафов и данных Excel.
                var groups = BuildGroups(d, s, report);
                var jobs = s.Cables ? NonModelExcel.Read(s.File) : new List<NonModelJob>();

                foreach (var j in jobs)
                    j.Cabinet = Group(j.Cabinet, groups);


                // Проверка данных или выполнение выбранных этапов.
                if (s.Preview)
                {
                    if (s.Cables)
                        Cables(d, s, groups, jobs, report, false);

                    if (s.Fasteners)
                        Fasteners(d, s, groups, report, false);

                    if (s.Copy)
                        Copy(d, s, groups, report, false);

                    if (s.Cables && (s.Fasteners || s.Copy))
                        report.Lines.Add("В проверке последующие этапы читают существующую модель; запланированные новые кабели/трубы ещё не созданы.");
                }
                else
                {
                    using (var group = new TransactionGroup(d, "ASML: немоделируемые элементы"))
                    {
                        group.Start();

                        try
                        {
                            if (s.Cables)
                                Stage(d, "Кабели и трубы", () => Cables(d, s, groups, jobs, report, true));

                            if (s.Fasteners)
                                Stage(d, "Крепления ПВХ", () => Fasteners(d, s, groups, report, true));

                            if (s.Copy)
                                Stage(d, "Секция и этаж", () => Copy(d, s, groups, report, true));

                            if (group.Assimilate() != TransactionStatus.Committed)
                                throw new InvalidOperationException("Группа изменений не подтверждена.");
                        }
                        catch (Exception ex)
                        {
                            if (group.GetStatus() == TransactionStatus.Started)
                                group.RollBack();

                            report.Changed = 0;
                            report.Lines.Insert(0, "ВСЕ ИЗМЕНЕНИЯ ЗАПУСКА ОТМЕНЕНЫ: " + ex.Message);
                            report.Show(c.Application, "Немоделируемые элементы — отмена");
                            return Result.Cancelled;
                        }
                    }
                }


                // Общий результат обработки.
                report.Show(c.Application, "Немоделируемые элементы — общий отчёт");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Util.Fail(ex, ref message);
            }
        }


        private static void Stage(Document d, string name, Action run)
        {
            using (var t = new Autodesk.Revit.DB.Transaction(d, "ASML: " + name))
            {
                t.Start();
                run();
                Util.Commit(t);
            }
        }


        private static string NormalCabinet(string value)
        {
            string text = Regex.Replace(value ?? "", @"\s+", "");
            var m = Regex.Match(text, @"^[СсCc](?<s>[0-9]+)ЩК(?<t>[0-9]+(?:\.[0-9]+)*)$", RegexOptions.IgnoreCase);
            return m.Success ? "С" + m.Groups["s"].Value + "ЩК" + m.Groups["t"].Value : text;
        }


        private static string Group(string value, Dictionary<string, string> groups)
        {
            string key = NormalCabinet(value);
            return groups.ContainsKey(key) ? groups[key] : key;
        }


        private static Dictionary<string, string> BuildGroups(Document d, NonModelSettings s, Report report)
        {
            string pattern = Regex.Escape(s.SumTemplate).Replace(Regex.Escape("{section}"), @"(?<section>[СсCc][0-9]+)").Replace(Regex.Escape("{type}"), @"(?<type>[0-9]+(?:\.[0-9]+)*)");
            var regex = new Regex("^" + pattern + "$");
            var groups = new Dictionary<string, string>();

            foreach (var gp in Util.Globals(d).Values)
            {
                var match = regex.Match(gp.Name);

                if (!match.Success)
                    continue;

                string section = match.Groups["section"].Value;

                if (section.Length > 0)
                    section = "С" + section.Substring(1);

                var rules = new FinishRules
                {
                    BaseTemplate = s.BaseTemplate,
                    SumTemplate = s.SumTemplate,
                    UseSections = s.BaseTemplate.Contains("{section}")
                };
                List<string> types;

                try
                {
                    types = rules.FormulaTypes(gp.GetFormula(), section);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Группа " + gp.Name + ": " + ex.Message);
                }

                string canonical = section + "ЩК" + match.Groups["type"].Value;

                foreach (string type in types)
                {
                    string key = section + "ЩК" + type;

                    if (groups.ContainsKey(key) && groups[key] != canonical)
                        throw new InvalidOperationException("Тип " + key + " входит в разные группы: " + groups[key] + " / " + canonical);

                    groups[key] = canonical;
                }

                // Представитель группы также соответствует собственному коду.
                if (groups.ContainsKey(canonical) && groups[canonical] != canonical)
                    throw new InvalidOperationException("Неоднозначный представитель группы: " + canonical);

                groups[canonical] = canonical;
                report.Lines.Add("Группа " + gp.Name + ": " + string.Join(", ", types.Select(t => section + "ЩК" + t)) + " → " + canonical);
            }

            report.Lines.Add("Ссылок группировки: " + groups.Count + ". Ключи без найденной группы используются точно.");
            return groups;
        }


        private static List<FamilyInstance> Instances(Document d, string family)
        {
            return new FilteredElementCollector(d).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(e => e.Symbol.FamilyName == family && e.SuperComponent == null).OrderBy(e => e.Id.IntegerValue).ToList();
        }


        private static Dictionary<string, FamilySymbol> Symbols(Document d, string family)
        {
            return new FilteredElementCollector(d).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.FamilyName == family).ToDictionary(x => x.Name, StringComparer.Ordinal);
        }


        private static XYZ Point(Element e)
        {
            if (e.Location is LocationPoint)
                return ((LocationPoint)e.Location).Point;

            if (e.Location is LocationCurve)
                return ((LocationCurve)e.Location).Curve.Evaluate(0.5, true);

            var b = e.get_BoundingBox(null);
            return b == null ? null : (b.Min + b.Max) / 2;
        }


        private static double M(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_METERS);
        }


        private static FamilyInstance Create(Document d, FamilySymbol sym, XYZ p, Level level)
        {
            if (!sym.IsActive)
            {
                sym.Activate();
                d.Regenerate();
            }

            if (sym.Family.FamilyPlacementType == FamilyPlacementType.ViewBased)
                return d.Create.NewFamilyInstance(p, sym, d.ActiveView);

            if (sym.Family.FamilyPlacementType == FamilyPlacementType.OneLevelBased || sym.Family.FamilyPlacementType == FamilyPlacementType.TwoLevelsBased)
                return d.Create.NewFamilyInstance(p, sym, level, StructuralType.NonStructural);

            return d.Create.NewFamilyInstance(p, sym, StructuralType.NonStructural);
        }


        private static void SetEstimate(Element e, string name, bool value)
        {
            var p = Util.Unique(e, name);

            if (p == null || p.IsReadOnly)
                throw new InvalidOperationException("Нет доступного параметра сметы «" + name + "».");

            bool ok;

            if (p.StorageType == StorageType.Integer)
                ok = p.Set(value ? 1 : 0);
            else if (p.StorageType == StorageType.String)
                ok = p.Set(value ? "Да" : "Нет");
            else
                throw new InvalidOperationException("Параметр сметы должен быть Да/Нет или текстом.");

            if (!ok)
                throw new InvalidOperationException("Запись сметы отклонена.");
        }


        private static void CopyEstimate(Document d, Element source, Element target, string name)
        {
            var p = Util.InstanceOrType(d, source, name);

            if (p == null || !p.HasValue)
                throw new InvalidOperationException("У трубы не заполнено «" + name + "».");

            bool value;

            if (p.StorageType == StorageType.Integer)
                value = p.AsInteger() != 0;
            else
            {
                string text = Util.Text(d, p).ToLowerInvariant();

                if (!new[]
                {
                    "да",
                    "нет",
                    "yes",
                    "no",
                    "true",
                    "false",
                    "0",
                    "1",
                    "истина",
                    "ложь"
                }.Contains(text))
                    throw new InvalidOperationException("Не распознано значение сметы: " + text);

                value = new[]
                {
                    "да",
                    "yes",
                    "true",
                    "1",
                    "истина"
                }.Contains(text);
            }

            SetEstimate(target, name, value);
        }


        private static void Clear(Element e, string name)
        {
            var p = Util.Unique(e, name);

            if (p == null || p.IsReadOnly)
                throw new InvalidOperationException("Нет доступного параметра «" + name + "».");

            bool ok;

            switch (p.StorageType)
            {
                case StorageType.String:
                    ok = p.Set("");
                    break;
                case StorageType.Integer:
                    ok = p.Set(0);
                    break;
                case StorageType.Double:
                    ok = p.Set(0.0);
                    break;
                case StorageType.ElementId:
                    ok = p.Set(ElementId.InvalidElementId);
                    break;
                default:
                    throw new InvalidOperationException("Неподдерживаемый StorageType при очистке.");
            }

            if (!ok)
                throw new InvalidOperationException("Очистка отклонена.");
        }


        private static XYZ Clamp(Document d, XYZ point, double margin)
        {
            var box = d.ActiveView.CropBox;

            if (box == null)
                return point;

            var local = box.Transform.Inverse.OfPoint(point);
            double minX = box.Min.X + margin, maxX = box.Max.X - margin, minY = box.Min.Y + margin, maxY = box.Max.Y - margin;

            if (minX > maxX)
                minX = maxX = (box.Min.X + box.Max.X) / 2;

            if (minY > maxY)
                minY = maxY = (box.Min.Y + box.Max.Y) / 2;

            return box.Transform.OfPoint(new XYZ(Math.Max(minX, Math.Min(maxX, local.X)), Math.Max(minY, Math.Min(maxY, local.Y)), local.Z));
        }


        private static void SetText(Document d, Element e, string name, string value)
        {
            var p = Util.Unique(e, name);

            if (p == null || p.IsReadOnly)
                throw new InvalidOperationException("Id " + e.Id.IntegerValue + ": нет доступного параметра " + name);

            bool ok;

            if (p.StorageType == StorageType.String)
                ok = p.Set(value);
            else if (p.StorageType == StorageType.Integer)
            {
                int n;
                ok = int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? p.Set(n) : p.SetValueString(value);
            }
            else
                ok = p.SetValueString(value);

            if (!ok)
                throw new InvalidOperationException("Запись " + name + " отклонена.");
        }


        private static void SetNumber(Element e, string name, double value)
        {
            var p = Util.Unique(e, name);

            if (p == null || p.IsReadOnly)
                throw new InvalidOperationException("Нет доступного параметра " + name + " у Id " + e.Id.IntegerValue);

            bool ok;

            if (p.StorageType == StorageType.Double)
            {
                double v = p.Definition.ParameterType == ParameterType.Length ? M(value) : value;
                ok = p.Set(v);
            }
            else if (p.StorageType == StorageType.Integer)
            {
                if (value != Math.Truncate(value) || value > int.MaxValue)
                    throw new InvalidOperationException(name + ": дробное/слишком большое значение для Integer");

                ok = p.Set((int)value);
            }
            else if (p.StorageType == StorageType.String)
                ok = p.Set(value.ToString(CultureInfo.InvariantCulture));
            else
                throw new InvalidOperationException("Неподходящий StorageType для " + name);

            if (!ok)
                throw new InvalidOperationException("Запись " + name + " отклонена.");
        }


        private static double Number(Document d, Element e, string name)
        {
            var p = Util.InstanceOrType(d, e, name);

            if (p == null || !p.HasValue)
                throw new InvalidOperationException("Не заполнено " + name + " у Id " + e.Id.IntegerValue);

            if (p.StorageType == StorageType.Double)
                return p.Definition.ParameterType == ParameterType.Length ? UnitUtils.ConvertFromInternalUnits(p.AsDouble(), DisplayUnitType.DUT_METERS) : p.AsDouble();

            if (p.StorageType == StorageType.Integer)
                return p.AsInteger();

            double n;

            if (!double.TryParse((p.AsString() ?? "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out n) || double.IsInfinity(n) || double.IsNaN(n))
                throw new InvalidOperationException("Некорректное число " + name);

            return n;
        }


        private static string TypeKey(string name)
        {
            string from = "ABCEHKMOPTXY", to = "АВСЕНКМОРТХУ";
            string value = name.ToUpperInvariant();

            for (int i = 0; i < from.Length; i++)
                value = value.Replace(from[i], to[i]);

            value = Regex.Replace(value, @"[Ø⌀∅]", "");
            value = Regex.Replace(value, @"(?<=\D)D(?=\s*\d)", "");
            value = Regex.Replace(value, @"(\d)\s*(ММ\.?)", "$1");
            return Regex.Replace(value, @"[\s\-–—._/\\]+", "_").Trim('_');
        }


        private static FamilySymbol Resolve(Dictionary<string, FamilySymbol> symbols, string name)
        {
            if (symbols.ContainsKey(name))
                return symbols[name];

            string key = TypeKey(name);
            var list = symbols.Values.Where(x => TypeKey(x.Name) == key || TypeKey(x.Name).Replace("ГОФРА", "ТРУБА") == key.Replace("ГОФРА", "ТРУБА")).Distinct().ToList();

            if (list.Count > 1)
                throw new InvalidOperationException("Неоднозначный тип: " + name);

            return list.SingleOrDefault();
        }


        private static XYZ FreeOrigin(Document d, NonModelSettings s)
        {
            var boxes = new List<BoundingBoxXYZ>();

            foreach (var e in new FilteredElementCollector(d).WhereElementIsNotElementType())
            {
                if (e is FamilyInstance && ((FamilyInstance)e).Symbol.FamilyName == s.TargetFamily)
                    continue;

                try
                {
                    var b = e.get_BoundingBox(null);

                    if (b != null && (e.LevelId == s.Level.Id || b.Min.Z - M(0.05) <= s.Level.Elevation && b.Max.Z + M(0.05) >= s.Level.Elevation))
                        boxes.Add(b);
                }
                catch
                {
                }
            }

            return boxes.Count == 0 ? new XYZ(0, M(s.Offset), s.Level.Elevation) : new XYZ((boxes.Min(b => b.Min.X) + boxes.Max(b => b.Max.X)) / 2, boxes.Max(b => b.Max.Y) + M(s.Offset), s.Level.Elevation);
        }


        private static void Cables(Document d, NonModelSettings s, Dictionary<string, string> groups, List<NonModelJob> jobs, Report report, bool write)
        {
            report.Lines.Add("=== КАБЕЛИ И ТРУБЫ: строк Excel " + jobs.Count + " ===");
            var symbols = Symbols(d, s.TargetFamily);

            if (symbols.Count == 0)
                throw new InvalidOperationException("Не загружено семейство " + s.TargetFamily);

            var cabinets = Instances(d, s.CabinetFamily).Where(e => Util.Source(d, e, s.Cabinet).Length > 0).GroupBy(e => Group(Util.Source(d, e, s.Cabinet), groups)).ToDictionary(g => g.Key, g => g.ToList());
            var pool = Instances(d, s.TargetFamily).GroupBy(e => Group(Util.Text(d, Util.Unique(e, s.Cabinet)), groups) + "\u001f" + e.Symbol.Name).ToDictionary(g => g.Key, g => new Queue<FamilyInstance>(g));
            var origin = FreeOrigin(d, s);
            double spacing = M(s.Spacing / 1000);
            var codes = jobs.Select(j => j.Cabinet).Distinct().OrderBy(x => x).ToList();
            var counters = new Dictionary<string, int>();
            var touched = new HashSet<string>();
            int created = 0, updated = 0, deleted = 0;

            foreach (var job in jobs.OrderBy(j => j.Cabinet).ThenBy(j => j.Type).ThenBy(j => !j.Estimate).ThenBy(j => j.SourceType))
            {
                var symbol = Resolve(symbols, job.Type);

                if (symbol == null)
                {
                    if (!s.DuplicateTypes)
                        throw new InvalidOperationException("Не найден тип " + job.Type);

                    var donors = symbols.Values.Where(x => TypeKey(x.Name).Contains("ТРУБА") || TypeKey(x.Name).Contains("ГОФРА")).OrderBy(x => x.Id.IntegerValue).ToList();
                    var donor = (TypeKey(job.Type).Contains("ТРУБА") || TypeKey(job.Type).Contains("ГОФРА") ? donors.FirstOrDefault() : null) ?? symbols.Values.OrderBy(x => x.Id.IntegerValue).First();
                    report.Lines.Add("Тип " + job.Type + ": дублирование «" + donor.Name + "»; параметры типа наследуются от донора.");

                    if (write)
                    {
                        symbol = (FamilySymbol)donor.Duplicate(job.Type);
                        symbols.Add(job.Type, symbol);
                    }
                    else
                        symbol = donor;
                }

                string actual = write ? symbol.Name : (symbols.ContainsKey(job.Type) ? job.Type : symbol.Name);
                string key = job.Cabinet + "\u001f" + actual;
                touched.Add(key);
                Queue<FamilyInstance> queue;
                pool.TryGetValue(key, out queue);
                var instance = queue != null && queue.Count > 0 ? queue.Dequeue() : null;
                bool near = job.Estimate && cabinets.ContainsKey(job.Cabinet);
                var cabinet = near ? cabinets[job.Cabinet].First() : null;
                var level = near ? d.GetElement(cabinet.LevelId) as Level : null;
                level = level ?? s.Level;

                if (near && cabinets[job.Cabinet].Count > 1)
                    report.Lines.Add("Шкаф " + job.Cabinet + ": несколько щитов, размещение у Id " + cabinet.Id.IntegerValue);

                int codeIndex = codes.IndexOf(job.Cabinet);
                XYZ basePoint = near ? Point(cabinet) : new XYZ(origin.X + (codeIndex % 8) * spacing * 7, origin.Y - (codeIndex / 8) * spacing * 12, origin.Z);

                if (basePoint == null)
                    throw new InvalidOperationException("Нет точки шкафа " + job.Cabinet);

                string zone = job.Cabinet + (near ? ":model" : ":free");
                int index = counters.ContainsKey(zone) ? counters[zone] : 0;
                counters[zone] = index + 1;
                var point = new XYZ(basePoint.X + spacing * (index % 5 + 1), basePoint.Y - spacing * (index / 5), basePoint.Z);

                try
                {
                    var before = point;
                    point = Clamp(d, point, spacing);

                    if (before.DistanceTo(point) > 0.0001)
                        report.Lines.Add("Шкаф " + job.Cabinet + ": точка ограничена границами вида.");
                }
                catch
                {
                    report.Lines.Add("Границы активного вида недоступны; точка размещения сохранена.");
                }

                string action = instance == null ? "Создать" : "Обновить";
                report.Lines.Add(job.Sheet + ", строка " + job.Row + ": " + action + " " + job.Cabinet + " / " + job.Type + "; количество=" + job.Quantity + "; СМ_Смета=" + job.Estimate + (instance == null ? "" : "; Id " + instance.Id.IntegerValue));

                if (!write)
                    continue;

                if (instance == null)
                {
                    instance = Create(d, symbol, point, level);
                    created++;
                }
                else
                {
                    updated++;

                    if (s.Move)
                    {
                        var oldPoint = Point(instance);

                        if (oldPoint == null)
                            throw new InvalidOperationException("Нет точки существующего экземпляра");

                        ElementTransformUtils.MoveElement(d, instance.Id, point - oldPoint);
                    }
                }

                SetText(d, instance, s.Cabinet, job.Cabinet);
                SetNumber(instance, s.Quantity, job.Quantity);
                SetEstimate(instance, s.Estimate, job.Estimate);

                if (s.ClearType)
                    Clear(instance, "ASML_Тип");

                report.Changed++;
            }

            if (s.DeleteExcess)
                foreach (var key in touched)
                    if (pool.ContainsKey(key))
                        foreach (var extra in pool[key])
                        {
                            report.Lines.Add((write ? "Удалён" : "Будет удалён") + " лишний кабель/труба Id " + extra.Id.IntegerValue);

                            if (write)
                            {
                                d.Delete(extra.Id);
                                deleted++;
                                report.Changed++;
                            }
                        }

            report.Lines.Add("Кабели/трубы: создано " + created + ", обновлено " + updated + ", удалено лишних " + deleted + ".");
        }


        private static int Diameter(FamilyInstance e)
        {
            foreach (string candidate in new[]
            {
                e.Symbol.Name,
                e.Symbol.FamilyName
            })
            {
                var m = Regex.Match(candidate, @"труба_?\s*пвх_?\s*(16|20|25|32)\s*$", RegexOptions.IgnoreCase);

                if (m.Success)
                    return int.Parse(m.Groups[1].Value);
            }

            int diameter;
            return int.TryParse(e.Symbol.Name, out diameter) && new[]
            {
                16,
                20,
                25,
                32
            }.Contains(diameter) && e.Symbol.FamilyName.IndexOf("ПВХ", StringComparison.OrdinalIgnoreCase) >= 0 && e.Symbol.FamilyName.IndexOf("Труба", StringComparison.OrdinalIgnoreCase) >= 0 ? diameter : 0;
        }


        private static void Fasteners(Document d, NonModelSettings s, Dictionary<string, string> groups, Report report, bool write)
        {
            report.Lines.Add("=== КРЕПЛЕНИЯ ПВХ; шаг " + s.Step + " м ===");
            var instances = new FilteredElementCollector(d).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(e => e.SuperComponent == null).OrderBy(e => e.Id.IntegerValue).ToList();
            var pipes = instances.Where(e => Diameter(e) > 0).ToList();

            if (pipes.Count == 0)
            {
                report.Lines.Add("Трубы ПВХ 16/20/25/32 не найдены; существующие крепления сохранены.");
                return;
            }

            var symbols = Symbols(d, s.TargetFamily);
            var diameters = pipes.Select(Diameter).Distinct();

            foreach (int diameter in diameters)
                if (!symbols.ContainsKey("Крепление_" + diameter))
                    throw new InvalidOperationException("Не найден тип Крепление_" + diameter + " в семействе " + s.TargetFamily);

            var old = Instances(d, s.TargetFamily).Where(e => new[] { "Крепление_16", "Крепление_20", "Крепление_25", "Крепление_32" }.Contains(e.Symbol.Name)).ToList();
            // План полностью проверяется перед удалением существующих экземпляров.
            var plan = new List<Tuple<FamilyInstance, int, double, XYZ, Level>>();

            foreach (var pipe in pipes)
            {
                double length = Number(d, pipe, s.Quantity);
                double count = Math.Ceiling(length / s.Step);

                if (length <= 0)
                {
                    report.Skipped++;
                    report.Lines.Add("Id " + pipe.Id.IntegerValue + ": длина <=0, пропущено");
                    continue;
                }

                if (count > int.MaxValue || double.IsInfinity(count) || double.IsNaN(count))
                    throw new InvalidOperationException("Некорректное количество креплений");

                var point = Point(pipe);

                if (point == null)
                    throw new InvalidOperationException("Нет точки трубы Id " + pipe.Id.IntegerValue);

                var level = d.GetElement(pipe.LevelId) as Level ?? s.Level;
                plan.Add(Tuple.Create(pipe, Diameter(pipe), count, point, level));
            }

            report.Lines.Add("Существующих креплений к пересозданию: " + old.Count + "; новых экземпляров: " + plan.Count);

            if (write)
            {
                foreach (var instance in old)
                    d.Delete(instance.Id);

                d.Regenerate();
                report.Changed += old.Count;
            }

            foreach (var item in plan)
            {
                report.Lines.Add("Труба Id " + item.Item1.Id.IntegerValue + " → Крепление_" + item.Item2 + ", количество=" + item.Item3);

                if (!write)
                    continue;

                var instance = Create(d, symbols["Крепление_" + item.Item2], item.Item4, item.Item5);
                SetNumber(instance, s.Quantity, item.Item3);
                SetText(d, instance, s.Cabinet, Group(Util.Source(d, item.Item1, s.Cabinet), groups));
                CopyEstimate(d, item.Item1, instance, s.Estimate);

                foreach (var bip in new[]
                {
                    BuiltInParameter.PHASE_CREATED,
                    BuiltInParameter.PHASE_DEMOLISHED,
                    BuiltInParameter.ELEM_PARTITION_PARAM
                })
                {
                    var source = item.Item1.get_Parameter(bip);
                    var target = instance.get_Parameter(bip);

                    if (source != null && target != null && !target.IsReadOnly)
                    {
                        if (source.StorageType == StorageType.ElementId)
                            target.Set(source.AsElementId());
                        else if (source.StorageType == StorageType.Integer)
                            target.Set(source.AsInteger());
                    }
                }

                report.Changed++;
            }
        }


        private static void Copy(Document d, NonModelSettings s, Dictionary<string, string> groups, Report report, bool write)
        {
            report.Lines.Add("=== КОПИРОВАНИЕ СЕКЦИИ И ЭТАЖА ===");
            var sources = Instances(d, s.CabinetFamily).GroupBy(e => Group(Util.Source(d, e, s.Cabinet), groups)).Where(g => g.Key.Length > 0).ToDictionary(g => g.Key, g => g.ToList());

            foreach (var target in Instances(d, s.TargetFamily))
            {
                string key = Group(Util.Source(d, target, s.Cabinet), groups);

                if (key.Length == 0 || !sources.ContainsKey(key))
                {
                    report.Skipped++;
                    report.Lines.Add("Id " + target.Id.IntegerValue + ": нет щита по ключу «" + key + "»");
                    continue;
                }

                var pool = sources[key];
                var payloads = pool.Select(e => Util.Source(d, e, s.Section) + "\u001f" + Util.Source(d, e, s.Floor)).Distinct().ToList();

                if (payloads.Count > 1)
                {
                    report.Skipped++;
                    report.Lines.Add("Id " + target.Id.IntegerValue + ": конфликт секции/этажа щитов по группе " + key + "; источники Id " + string.Join(", ", pool.Select(e => e.Id.IntegerValue)) + ". Копирование пропущено.");
                    continue;
                }

                var source = pool.First();

                foreach (string name in new[]
                {
                    s.Section,
                    s.Floor
                })
                {
                    string value = Util.Source(d, source, name), oldValue = Util.Text(d, Util.Unique(target, name));

                    if (value.Length == 0 || (!s.Overwrite && oldValue.Length > 0) || value == oldValue)
                    {
                        report.Skipped++;
                        report.Lines.Add("Id " + target.Id.IntegerValue + ": " + name + " без изменений; источник=«" + value + "», приёмник=«" + oldValue + "»");
                        continue;
                    }

                    report.Lines.Add("Id " + target.Id.IntegerValue + ": " + name + " «" + oldValue + "» → «" + value + "»; источник Id " + source.Id.IntegerValue);

                    if (write)
                    {
                        SetText(d, target, name, value);
                        report.Changed++;
                    }
                }
            }
        }
    }
}

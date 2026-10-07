using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [Transaction(TransactionMode.Manual)]
    public class FinishCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string m, ElementSet e)
        {
            return Finish.Run(c, ref m);
        }
    }


    internal static class Finish
    {
        private sealed class Write
        {
            internal Parameter P;
            internal string Value;
            internal GlobalParameter Global;
            internal int ElementId;
            internal string Key
            {
                get
                {
                    return P.Element.Id.IntegerValue + ":" + P.Id.IntegerValue;
                }
            }
        }


        internal static Result Run(ExternalCommandData c, ref string message)
        {
            try
            {
                // Проверка проекта до открытия настроек и изменения модели.
                if (!Util.IsSetunProject(c))
                    return Result.Cancelled;

                var d = Util.Project(c);
                var globals = Util.Globals(d);
                var report = new Report();
                FinishRules settings;


                // Выбор правил определения отделки.
                using (var form = new FinishRulesForm(d, FinishRules.Load(d)))
                {
                    if (form.ShowDialog(new WindowOwner(c.Application)) != WinForms.DialogResult.OK)
                        return Result.Cancelled;

                    settings = form.Settings;
                }

                ScopeSettings scope;

                using (var form = new ScopeForm(d, "Unified"))
                {
                    if (form.ShowDialog(new WindowOwner(c.Application)) != WinForms.DialogResult.OK)
                        return Result.Cancelled;

                    scope = form.Settings;
                }


                // Подготовка записей по выбранной области модели.
                var regex = settings.Pattern();
                var writes = new List<Write>();

                foreach (var element in new FilteredElementCollector(d).WhereElementIsNotElementType().WhereElementIsViewIndependent().ToElements())
                {
                    try
                    {
                        if (!scope.Matches(d, element))
                            continue;

                        var source = Util.Unique(element, settings.Source);

                        if (source == null || !source.HasValue)
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + element.Id.IntegerValue + ": исходный параметр отсутствует или пуст.");
                            continue;
                        }

                        var family = (element as FamilyInstance)?.Symbol.FamilyName ?? (d.GetElement(element.GetTypeId()) as ElementType)?.FamilyName ?? "";

                        if (!string.IsNullOrWhiteSpace(settings.ExcludeFamily) && family.IndexOf(settings.ExcludeFamily.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + element.Id.IntegerValue + ": исключён по семейству.");
                            continue;
                        }

                        var match = regex.Match(Util.Text(d, source));

                        if (!match.Success)
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + element.Id.IntegerValue + ": не распознан шкаф.");
                            continue;
                        }

                        string suffix = match.Groups["type"].Value;
                        string section = settings.UseSections ? settings.SectionMarker + match.Groups["section"].Value : "";

                        if (settings.UseSections && settings.Sections().Count > 0 && !settings.Sections().Contains(section, StringComparer.OrdinalIgnoreCase))
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + element.Id.IntegerValue + ": секция " + section + " не включена в правила.");
                            continue;
                        }

                        string sumName = settings.GlobalName(settings.SumTemplate, section, suffix);
                        string baseName = settings.GlobalName(settings.BaseTemplate, section, suffix);
                        string name = settings.UseSum && globals.ContainsKey(sumName) ? sumName : settings.Fallback ? baseName : "";
                        GlobalParameter gp;

                        if (!globals.TryGetValue(name, out gp))
                            throw new InvalidOperationException("Не найден подходящий глобальный параметр: " + sumName + " / " + baseName);
                        // Состав суммы берём из формулы, без словаря и ручной таблицы.
                        string finish;

                        try
                        {
                            var types = name == sumName && settings.UseSum ? settings.FormulaTypes(gp.GetFormula(), section) : new List<string>
                            {
                                suffix
                            };
                            finish = settings.FinishText(section, types);
                        }
                        catch (Exception ex)
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + element.Id.IntegerValue + ": " + name + ": " + ex.Message + ". Записи коэффициента и текста пропущены.");
                            continue;
                        }

                        report.Lines.Add("Id " + element.Id.IntegerValue + ": шкаф «" + Util.Text(d, source) + "», глобальный параметр «" + name + "», текст «" + finish + "».");
                        // Независимые требования: конфликт коэффициента не блокирует текст и наоборот.
                        Add(d, element, settings.Target, name, gp, writes, report);
                        Add(d, element, settings.Finish, finish, null, writes, report);
                    }
                    catch (Exception ex)
                    {
                        report.Lines.Add("Id " + element.Id.IntegerValue + ": " + ex.Message);
                        report.Skipped++;
                    }
                }


                // Запись параметров с проверкой конфликтов.
                using (var t = new Autodesk.Revit.DB.Transaction(d, "ASML: тип отделки"))
                {
                    t.Start();

                    foreach (var group in writes.GroupBy(w => w.Key))
                    {
                        var w = group.First();

                        if (group.Select(x => (x.Global == null ? "text:" : "global:") + x.Value).Distinct().Count() > 1)
                        {
                            report.Skipped += group.Count();
                            report.Lines.Add("Конфликт параметра владельца Id " + w.P.Element.Id.IntegerValue + ": " + string.Join("; ", group.Select(x => "Id " + x.ElementId + " → " + x.Value)) + ". Запись пропущена.");
                            continue;
                        }

                        try
                        {
                            if (w.P.IsReadOnly)
                                throw new InvalidOperationException("Параметр только для чтения.");

                            if (w.Global != null)
                            {
                                if (w.P.GetAssociatedGlobalParameter() == w.Global.Id)
                                {
                                    report.Skipped++;
                                    report.Lines.Add("Id " + w.ElementId + ": связь уже задана — " + w.Value);
                                    continue;
                                }

                                if (!w.P.CanBeAssociatedWithGlobalParameter(w.Global.Id))
                                    throw new InvalidOperationException("Типы параметров несовместимы или связь недоступна.");
                            }
                            else
                            {
                                if (w.P.StorageType != StorageType.String)
                                    throw new InvalidOperationException("Тип отделки должен быть текстовым параметром.");

                                if ((w.P.AsString() ?? "") == w.Value)
                                {
                                    report.Skipped++;
                                    report.Lines.Add("Id " + w.ElementId + ": текст уже совпадает.");
                                    continue;
                                }
                            }

                            using (var sub = new SubTransaction(d))
                            {
                                sub.Start();

                                if (w.Global != null)
                                {
                                    if (w.P.GetAssociatedGlobalParameter() != ElementId.InvalidElementId)
                                        w.P.DissociateFromGlobalParameter();

                                    w.P.AssociateWithGlobalParameter(w.Global.Id);
                                }
                                else if (!w.P.Set(w.Value))
                                    throw new InvalidOperationException("Запись отклонена.");

                                if (sub.Commit() != TransactionStatus.Committed)
                                    throw new InvalidOperationException("Запись не подтверждена.");
                            }

                            report.Changed++;
                            report.Lines.Add("Владелец Id " + w.P.Element.Id.IntegerValue + ": " + w.P.Definition.Name + " → " + w.Value);
                        }
                        catch (Exception ex)
                        {
                            report.Skipped++;
                            report.Lines.Add("Id " + w.ElementId + ": " + ex.Message);
                        }
                    }

                    Util.Commit(t);
                }


                // Сохранение настроек и вывод отчёта.
                if (settings.Remember)
                {
                    try
                    {
                        settings.Save(d);
                    }
                    catch (Exception ex)
                    {
                        report.Lines.Add("Настройки не сохранены: " + ex.Message);
                    }
                }

                report.Show(c.Application, "Тип отделки — отчёт");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Util.Fail(ex, ref message);
            }
        }


        private static void Add(Document d, Element element, string name, string value, GlobalParameter gp, List<Write> writes, Report report)
        {
            try
            {
                var p = Util.InstanceOrType(d, element, name);

                if (p == null)
                    throw new InvalidOperationException("Не найден параметр «" + name + "».");

                writes.Add(new Write { P = p, Value = value, Global = gp, ElementId = element.Id.IntegerValue });
            }
            catch (Exception ex)
            {
                report.Skipped++;
                report.Lines.Add("Id " + element.Id.IntegerValue + ": " + ex.Message);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [Transaction(TransactionMode.Manual)]
    public class SortCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                // Проверка проекта до открытия настроек и изменения модели.
                if (!Util.IsSetunProject(c))
                    return Result.Cancelled;

                var d = Util.Project(c);
                SortSettings settings;


                // Выбор исходных параметров и разделителя.
                using (var form = new SortForm(d, SortSettings.Load(d)))
                {
                    if (form.ShowDialog(new WindowOwner(c.Application)) != F.DialogResult.OK)
                        return Result.Cancelled;

                    settings = form.Settings;
                }

                var report = new SortReport();
                report.Notes.Add("Область: все экземпляры текущего документа. Настройки видов-спецификаций не изменяются.");


                // Формирование и запись ключей сортировки.
                using (var t = new Autodesk.Revit.DB.Transaction(d, "ASML: сортировка спецификации"))
                {
                    t.Start();

                    foreach (var e in new FilteredElementCollector(d).WhereElementIsNotElementType().ToElements())
                    {
                        var row = new SortRecord
                        {
                            Id = e.Id.IntegerValue,
                            Status = "Ошибка",
                            Category = "",
                            ElementName = "",
                            OldValue = "",
                            NewValue = "",
                            SourceValues = "",
                            Reason = ""
                        };
                        report.Rows.Add(row);

                        try
                        {
                            row.Category = e.Category?.Name ?? "Без категории";
                            row.ElementName = e.Name;
                            var ps = e.GetParameters(settings.Target).ToList();

                            if (ps.Count == 0)
                            {
                                row.Status = "Нет целевого параметра";
                                row.Reason = "У экземпляра отсутствует «" + settings.Target + "».";
                                continue;
                            }

                            var writable = ps.Where(p => !p.IsReadOnly && p.StorageType == StorageType.String).ToList();

                            if (writable.Count > 1)
                            {
                                row.Status = "Неоднозначный параметр";
                                row.Reason = "Несколько доступных текстовых параметров с целевым именем.";
                                continue;
                            }

                            Parameter target = writable.Count == 1 ? writable[0] : ps[0];
                            row.OldValue = Util.Text(d, target);

                            if (writable.Count == 0)
                            {
                                row.Status = ps.All(p => p.StorageType != StorageType.String) ? "Неверный тип параметра" : "Только для чтения";
                                row.Reason = string.Join("; ", ps.Select(p => "StorageType=" + p.StorageType + ", IsReadOnly=" + p.IsReadOnly));
                                continue;
                            }

                            var values = new List<string>();
                            var detail = new List<string>();

                            foreach (string name in settings.Sources)
                            {
                                string value = ReadSource(d, e, name, out string origin);
                                values.Add(value);
                                detail.Add(name + " [" + origin + "] = «" + value + "»");
                                row.SourceValues = string.Join("; ", detail);
                            }

                            string key = values.All(v => v.Length == 0) ? "" : string.Join(settings.Separator, values);
                            row.NewValue = key;

                            if ((target.AsString() ?? "") == key)
                            {
                                row.Status = key.Length == 0 ? "Нет исходных данных" : "Без изменений";
                                row.Reason = key.Length == 0 ? "Все исходные значения пусты; целевой параметр уже пуст." : "Старый ключ совпадает с рассчитанным.";
                                continue;
                            }

                            using (var sub = new SubTransaction(d))
                            {
                                sub.Start();

                                if (!target.Set(key))
                                    throw new InvalidOperationException("Запись отклонена.");

                                if (sub.Commit() != TransactionStatus.Committed)
                                    throw new InvalidOperationException("Запись не подтверждена Revit.");
                            }

                            row.Status = key.Length == 0 ? "Очищено" : "Изменено";
                            row.Reason = key.Length == 0 ? "Все исходные значения пусты; старый ключ очищен." : values.Any(v => v.Length == 0) ? "Ключ записан; пустые части сохранены между разделителями." : "Ключ записан.";
                        }
                        catch (Exception ex)
                        {
                            row.Status = "Ошибка";
                            row.Reason = ex.Message;
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
                        report.Notes.Add("Настройки не сохранены: " + ex.Message);
                    }
                }

                report.Show(c.Application, settings, d.Title);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Util.Fail(ex, ref message);
            }
        }


        private static string ReadSource(Document d, Element e, string name, out string origin)
        {
            var p = Util.Unique(e, name);
            string value = Util.Text(d, p);

            if (value.Length > 0)
            {
                origin = "экземпляр";
                return value;
            }

            var type = d.GetElement(e.GetTypeId());
            var tp = Util.Unique(type, name);
            value = Util.Text(d, tp);

            if (value.Length > 0)
            {
                origin = "тип Id " + type.Id.IntegerValue;
                return value;
            }

            origin = p == null && tp == null ? "параметр отсутствует" : "значение пусто";
            return "";
        }
    }
}

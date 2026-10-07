using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [Transaction(TransactionMode.Manual)]
    public class GlobalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                // Проверка проекта до открытия настроек и изменения модели.
                if (!Util.IsSetunProject(c))
                    return Result.Cancelled;

                var d = Util.Project(c);
                string path;
                bool update;


                // Выбор Excel-файла и режима обновления.
                using (var form = new GlobalForm())
                {
                    if (form.ShowDialog(new WindowOwner(c.Application)) != WinForms.DialogResult.OK)
                        return Result.Cancelled;

                    path = form.Path;
                    update = form.UpdateExisting;
                }


                // Чтение данных и подготовка отчёта.
                var rows = Xlsx.Read(path);
                var globals = Util.Globals(d);
                var skip = new HashSet<string>();
                var report = new Report();
                // Один импорт — одна атомарная транзакция: ошибка откатывает весь импорт.
                using (var t = new Autodesk.Revit.DB.Transaction(d, "ASML: глобальные параметры"))
                {
                    t.Start();

                    foreach (var row in rows)
                    {
                        if (globals.ContainsKey(row.Name))
                        {
                            if (!update)
                            {
                                skip.Add(row.Name);
                                report.Skipped++;
                                continue;
                            }

                            if (globals[row.Name].GetDefinition().ParameterType != ParameterType.Number)
                                throw new InvalidOperationException(row.Name + ": существующий параметр не типа Number.");
                        }
                        else
                            globals.Add(row.Name, GlobalParameter.Create(d, row.Name, ParameterType.Number));
                    }

                    d.Regenerate();


                    // Запись числовых значений.
                    foreach (var row in rows.Where(r => r.Value.HasValue && !skip.Contains(r.Name)))
                    {
                        var p = globals[row.Name];

                        if (p.IsDrivenByDimension || p.IsDrivenByFormula || p.IsReporting)
                            throw new InvalidOperationException(row.Name + ": значение управляется формулой или размером; импорт отменён.");

                        p.SetValue(new DoubleParameterValue(row.Value.Value));
                        report.Changed++;
                        report.Lines.Add(row.Name + " = " + row.Value);
                    }


                    // Назначение формул после создания всех параметров.
                    foreach (var row in rows.Where(r => r.Formula.Length > 0 && !skip.Contains(r.Name)))
                    {
                        var p = globals[row.Name];

                        if (p.IsDrivenByDimension || p.IsReporting || !p.IsValidFormula(row.Formula))
                            throw new InvalidOperationException(row.Name + ": Revit не принимает формулу «" + row.Formula + "».");

                        p.SetFormula(row.Formula);
                        report.Changed++;
                        report.Lines.Add(row.Name + " = " + row.Formula);
                    }

                    d.Regenerate();
                    Util.Commit(t);
                }


                // Результат импорта.
                report.Show(c.Application, "Глобальные параметры");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "Импорт не выполнен; изменения отменены.\n" + ex.Message;
                return Result.Failed;
            }
        }
    }
}

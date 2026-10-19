using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.ViewModels;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Parameters_Ribbon.Common.CopyElemParamData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.Command
{
    public class CommandWriteValues : IExecutableCommand
    {
        private readonly List<ParameterRuleElement> _rules;
        private readonly bool _isWriteInGroups;

        public CommandWriteValues(ObservableCollection<ParameterRuleElement> rules, bool isWriteInGroups)
        {
            _rules = rules.ToList();
            _isWriteInGroups = isWriteInGroups;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null)
                return Result.Cancelled;

            var progress = new ProgressInfoViewModel
            {
                ProcessTitle = "Копирование параметров",
                MaintStatus = "Собираю элементы",
                IsIndeterminate = true,
            };
            var window = ProgressWindow.ShowInBackground(progress, "KPLN: Копирование параметров");

            try
            {
                var elementsByRule = new List<Element[]>();
                foreach (ParameterRuleElement rule in _rules)
                {
                    progress.ThrowIfCancellationRequested();
                    elementsByRule.Add(new FilteredElementCollector(doc)
                        .OfCategoryId((rule.SelectedCategory.Data as Category).Id)
                        .WhereElementIsNotElementType()
                        .ToArray());
                }

                // При отмене откатываем все этапы, включая уже завершённую запись в группах.
                using (var group = new TransactionGroup(doc, "KPLN: Копирование параметров"))
                {
                    group.Start();
                    if (_isWriteInGroups)
                        WriteRules(doc, elementsByRule, progress, true);

                    WriteRules(doc, elementsByRule, progress, false);
                    progress.BeginCommit();
                    if (group.Assimilate() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Копирование параметров не сохранено.");
                }

                progress.IsComplete = true;
                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                TaskDialog.Show("KPLN: Копирование параметров", "Копирование отменено. Изменения не сохранены.");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return Result.Failed;
            }
            finally
            {
                window.CloseSafely();
            }
        }

        private void WriteRules(Document doc, List<Element[]> elementsByRule, ProgressInfoViewModel progress, bool writeAll)
        {
            string stage = writeAll ? "Заполняю параметры с учётом групп" : "Заполняю параметры вне групп и с разными значениями";
            int total = elementsByRule.Sum(elements => elements.Length);
            int completed = 0;
            progress.Report(0, total, stage);

            using (var transaction = new Transaction(doc, writeAll ? "KPLN: Коп пар-в: ВСЕ" : "KPLN: Коп пар-в: ВНЕ групп"))
            {
                transaction.Start();
                for (int index = 0; index < _rules.Count; index++)
                {
                    ParameterRuleElement rule = _rules[index];
                    foreach (Element element in elementsByRule[index])
                    {
                        progress.ThrowIfCancellationRequested();
                        Parameter sourceParameter = GetParameterByElement(element, rule.SelectedSourceParameter);
                        Parameter targetParameter = GetParameterByElement(element, rule.SelectedTargetParameter)
                            ?? throw new InvalidOperationException($"Не удалось получить параметр {rule.SelectedTargetParameter.Name} для элемента с id: {element.Id}");
                        if (writeAll && sourceParameter == null)
                            throw new InvalidOperationException($"Не удалось получить параметр {rule.SelectedSourceParameter.Name} для элемента с id: {element.Id}");

                        if (!(targetParameter.Definition is InternalDefinition definition))
                            throw new InvalidOperationException($"Не удалось преобразовать в InternalDefinition для параметра {targetParameter.Definition.Name}");

                        if (writeAll || definition.VariesAcrossGroups || element.GroupId.Equals(ElementId.InvalidElementId))
                            SetElemParamByRule(sourceParameter, targetParameter);

                        // Считаем каждый просмотренный элемент, в том числе пропущенный по условию группы.
                        progress.Report(++completed, total, stage);
                    }
                }

                progress.SetStage("Сохраняю результаты этапа");
                if (transaction.Commit() != TransactionStatus.Committed)
                    throw new InvalidOperationException("Не удалось сохранить этап копирования параметров.");
            }
        }

        private static Parameter GetParameterByElement(Element element, ListBoxElement rule)
        {
            foreach (Parameter p in element.Parameters)
            {
                if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                {
                    return p;
                }
            }

            Element elemType = element.Document.GetElement(element.GetTypeId());
            if (elemType != null)
            {
                foreach (Parameter p in elemType.Parameters)
                {
                    if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                    {
                        return p;
                    }
                }
            }

            return null;
        }

        private static void SetElemParamByRule(Parameter sourceParameter, Parameter targetParameter)
        {
            if (sourceParameter != null && targetParameter != null)
            {
                switch (targetParameter.StorageType)
                {
                    case StorageType.Double:
                        double? dv = GetDoubleValue(sourceParameter);
                        if (dv != null)
                        {
                            targetParameter.Set((double)dv);
                        }
                        break;
                    case StorageType.Integer:
                        int? iv = GetIntegerValue(sourceParameter);
                        if (iv != null)
                        {
                            targetParameter.Set((int)iv);
                        }
                        break;
                    case StorageType.String:
                        string sv = GetStringValue(sourceParameter);
                        if (sv != null && sv != " " && sv != string.Empty)
                        {
                            targetParameter.Set(sv);
                        }
                        break;
                }
            }
        }

        private static double? GetDoubleValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return p.AsDouble();
                    case StorageType.Integer:
                        return p.AsInteger();
                    case StorageType.String:
                        return double.Parse(p.AsString(), System.Globalization.NumberStyles.Float);
                    default:
                        return null;
                }
            }
            catch (Exception) { return null; }
        }

        private static int? GetIntegerValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return (int)Math.Round(p.AsDouble());
                    case StorageType.Integer:
                        return p.AsInteger();
                    case StorageType.String:
                        return int.Parse(p.AsString(), System.Globalization.NumberStyles.Integer);
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string GetStringValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return p.AsValueString();
                    case StorageType.Integer:
                        return p.AsValueString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                try
                {
                    switch (p.StorageType)
                    {
                        case StorageType.Double:
                            return p.AsDouble().ToString();
                        case StorageType.Integer:
                            return p.AsInteger().ToString();
                        case StorageType.String:
                            return p.AsString();
                        default:
                            return null;
                    }
                }
                catch (Exception) { return null; }
            }
        }
    }
}

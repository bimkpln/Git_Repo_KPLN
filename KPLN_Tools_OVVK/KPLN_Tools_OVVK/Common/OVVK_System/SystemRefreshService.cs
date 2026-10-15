using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal sealed class SystemRefreshService
    {
        private readonly Document _doc;
        internal int ChangedTypeCount { get; private set; }
        internal int RefreshedSystemCount { get; private set; }
        internal List<string> Warnings { get; } = new List<string>();

        internal SystemRefreshService(Document doc) { _doc = doc; }

        internal static string GetAbbreviation(string abbreviation, string typeName)
        {
            // Сохраняем префикс из Python-версии и завершающее подчёркивание.
            if (string.IsNullOrEmpty(abbreviation))
                return abbreviation;

            string normalized = abbreviation.EndsWith("_", StringComparison.Ordinal) ? abbreviation : abbreviation + "_";
            // Имя типа само может содержать "_": корректное окончание нельзя разбирать на отдельные части.
            if (string.Equals(normalized, typeName + "_", StringComparison.Ordinal)
                || normalized.EndsWith("_" + typeName + "_", StringComparison.Ordinal))
                return normalized;

            string[] parts = normalized.Split('_');
            string prefix = string.Join("_", parts.Take(parts.Length - 2));
            return (parts.Length > 2 ? prefix + "_" : string.Empty) + typeName + "_";
        }

        internal void Execute(Action<int, int, string> report, Action checkCancellation, Action beforeCommit)
        {
            MEPSystemType[] types = new FilteredElementCollector(_doc).OfClass(typeof(MechanicalSystemType))
                .Cast<MEPSystemType>().Concat(new FilteredElementCollector(_doc).OfClass(typeof(PipingSystemType))
                    .Cast<MEPSystemType>()).ToArray();
            var originalAbbreviations = types.ToDictionary(type => type.Id, type => type.Abbreviation);

            // По одному жёсткому участку на фактическую систему, без поиска по части имени.
            var representatives = new Dictionary<ElementId, ElementId>();
            var curves = new FilteredElementCollector(_doc).OfClass(typeof(Duct)).Cast<MEPCurve>()
                .Concat(new FilteredElementCollector(_doc).OfClass(typeof(Pipe)).Cast<MEPCurve>());
            foreach (MEPCurve curve in curves)
            {
                checkCancellation();
                MEPSystem system = curve.MEPSystem;
                if (system != null && !representatives.ContainsKey(system.Id))
                    representatives.Add(system.Id, curve.Id);
            }

            var curveTypes = representatives.Values.ToDictionary(id => id, id => ((MEPCurve)_doc.GetElement(id)).MEPSystem.GetTypeId());
            var temporaryTypes = new Dictionary<ElementId, ElementId>();
            using (var group = new TransactionGroup(_doc, "KPLN: Обновление имени системы"))
            {
                group.Start();
                using (var preparation = new Transaction(_doc, "KPLN: Подготовка типов систем"))
                {
                    preparation.Start();
                    int completed = 0;
                    foreach (MEPSystemType type in types)
                    {
                        report(completed++, types.Length, "Подготавливаю типы систем");
                        string abbreviation = originalAbbreviations[type.Id];
                        if (!string.IsNullOrEmpty(abbreviation) && !abbreviation.EndsWith("_", StringComparison.Ordinal))
                            type.Abbreviation = abbreviation + "_";
                    }

                    foreach (ElementId typeId in curveTypes.Values.Distinct())
                    {
                        checkCancellation();
                        // Копия сохраняет классификацию исходного типа, включая системы пожаротушения.
                        var source = (MEPSystemType)_doc.GetElement(typeId);
                        ElementType temporary = source.Duplicate("KPLN_refresh_" + Guid.NewGuid().ToString("N"));
                        temporaryTypes.Add(typeId, temporary.Id);
                    }
                    checkCancellation();
                    if (preparation.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Не удалось подготовить типы систем.");
                }


                using (var transaction = new Transaction(_doc, "KPLN: Обновление систем"))
                {
                    transaction.Start();
                    int completed = 0;
                    foreach (MEPSystemType type in types)
                    {
                        report(completed++, types.Length, "Сопоставляю сокращения с типами систем");
                        string original = originalAbbreviations[type.Id];
                        string desired = GetAbbreviation(original, type.Name);
                        if (type.Abbreviation != desired)
                            type.Abbreviation = desired;
                        if (original != desired)
                            ChangedTypeCount++;
                    }
                    _doc.Regenerate();

                    completed = 0;
                    foreach (var item in curveTypes)
                    {
                        report(completed++, curveTypes.Count, "Обновляю системы элементов");
                        try
                        {
                            using (var change = new SubTransaction(_doc))
                            {
                                change.Start();
                                var curve = _doc.GetElement(item.Key) as MEPCurve;
                                if (curve == null)
                                    throw new InvalidOperationException("Участок системы больше не существует.");
                                SetSystemType(curve, temporaryTypes[item.Value]);
                                _doc.Regenerate();
                                SetSystemType(curve, item.Value);
                                _doc.Regenerate();
                                if (curve.MEPSystem == null || !curve.MEPSystem.GetTypeId().Equals(item.Value))
                                    throw new InvalidOperationException("Исходный тип системы не восстановлен.");
                                if (change.Commit() != TransactionStatus.Committed)
                                    throw new InvalidOperationException("Обновление системы отменено.");
                            }
                            RefreshedSystemCount++;
                        }
                        catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            Warnings.Add("Элемент " + item.Key + ": " + ex.Message);
                        }
                    }

                    report(0, 1, "Удаляю временные типы систем");
                    var temporaryIds = new HashSet<ElementId>(temporaryTypes.Values);
                    var allowedDeletions = new HashSet<ElementId>(temporaryIds);
                    foreach (MEPSystem system in new FilteredElementCollector(_doc).OfClass(typeof(MechanicalSystem)).Cast<MEPSystem>()
                        .Concat(new FilteredElementCollector(_doc).OfClass(typeof(PipingSystem)).Cast<MEPSystem>()))
                    {
                        if (temporaryIds.Contains(system.GetTypeId()))
                        {
                            if (system.Elements.Size > 0 || system.BaseEquipment != null)
                                throw new InvalidOperationException("Во временной системе остались элементы. Обновление отменено.");
                            allowedDeletions.Add(system.Id);
                        }
                    }
                    foreach (ElementId id in temporaryIds)
                    {
                        checkCancellation();
                        ICollection<ElementId> deleted = _doc.Delete(id);
                        if (deleted.Any(deletedId => !allowedDeletions.Contains(deletedId)))
                            throw new InvalidOperationException("Удаление временного типа затронуло элементы модели. Обновление отменено.");
                    }

                    beforeCommit();
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Обновление систем не сохранено.");
                }
                if (group.Assimilate() != TransactionStatus.Committed)
                    throw new InvalidOperationException("Обновление систем отменено.");
            }
        }

        private static void SetSystemType(MEPCurve curve, ElementId typeId)
        {
            if (curve is Duct duct)
                duct.SetSystemType(typeId);
            else if (curve is Pipe pipe)
                pipe.SetSystemType(typeId);
            else
                throw new InvalidOperationException("Обновление поддерживается для труб и воздуховодов.");
        }
    }
}

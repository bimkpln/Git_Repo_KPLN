using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed partial class VentFamilyService
    {
        private sealed class TypeInput
        {
            internal FamilyParameter Parameter;
            internal object Value;
        }

        private sealed class PreservedType
        {
            internal FamilyType Type;
            internal string Name;
            internal List<TypeInput> Inputs;
        }

        private static object ReadTypeInput(FamilyType type, FamilyParameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    return type.AsDouble(parameter);
                case StorageType.Integer:
                    return type.AsInteger(parameter);
                case StorageType.String:
                    return type.AsString(parameter) ?? string.Empty;
                case StorageType.ElementId:
                    var id = type.AsElementId(parameter);
                    return id == null ? (object)null : Common.IDHelper.ElIdValue(id);
                default:
                    return null;
            }
        }

        private static List<TypeInput> CaptureTypeInputs(FamilyManager manager, FamilyType type)
        {
            var result = new List<TypeInput>();
            foreach (FamilyParameter parameter in manager.Parameters)
                if (!parameter.IsDeterminedByFormula && !parameter.IsReadOnly && !parameter.IsReporting)
                    result.Add(new TypeInput { Parameter = parameter, Value = ReadTypeInput(type, parameter) });
            return result;
        }

        private static void VerifyTypeInputs(FamilyType type, IEnumerable<TypeInput> inputs, ISet<string> allowed, string message)
        {
            var changed = new List<string>();
            foreach (var input in inputs)
            {
                if (allowed != null && allowed.Contains(FamilyParameterNames.Canonical(input.Parameter.Definition.Name)))
                    continue;
                object current = ReadTypeInput(type, input.Parameter);
                bool same = input.Value is double && current is double ? Math.Abs((double)input.Value - (double)current) < 1e-9 : Equals(input.Value, current);
                if (!same)
                    changed.Add(input.Parameter.Definition.Name + ": «" + Convert.ToString(input.Value, CultureInfo.InvariantCulture)
                    + "» → «" + Convert.ToString(current, CultureInfo.InvariantCulture) + "»");
            }
            if (changed.Count > 0)
                throw new InvalidOperationException(message + ":\n" + string.Join("\n", changed)
                + "\nЧисловые значения здесь приведены во внутренних единицах Revit. Изменения отменены.");
        }

        private static void EnsureWritable(FamilyParameter parameter)
        {
            if (parameter.IsDeterminedByFormula)
                throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "» вычисляется по формуле: " + parameter.Formula);
            if (parameter.IsReadOnly || parameter.IsReporting)
                throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "» доступен только для чтения.");
        }

        private static void SetIntegerValue(FamilyManager manager, string name, int value)
        {
            var parameter = manager.get_Parameter(name);
            if (parameter == null || parameter.StorageType != StorageType.Integer)
                throw new InvalidOperationException("В семействе нет целочисленного параметра «" + name + "».");
            // Существующее значение не требует записи, в том числе при формуле или блокировке параметра.
            if (manager.CurrentType.AsInteger(parameter) == value)
                return;
            // Формулы пересчитает Revit после записи независимых входов; результат проверяем перед сохранением.
            if (parameter.IsDeterminedByFormula)
                return;
            EnsureWritable(parameter);
            try
            {
                manager.Set(parameter, value);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Параметр «" + name + "», значение " + value + ": " + OperationError.Message(ex), ex);
            }
        }

        private static bool? ReadBooleanValue(FamilyManager manager, FamilyType type, string name)
        {
            var parameter = FindFamilyParameter(manager, name);
            if (parameter == null)
                return null;
            if (parameter.StorageType != StorageType.Integer)
                throw new InvalidOperationException("Параметр «" + name + "» должен быть логическим (Да/Нет).");
            int? value = type.AsInteger(parameter);
            if (value.HasValue && value != 0 && value != 1)
                throw new InvalidOperationException("Некорректное логическое значение «" + name + "».");
            return value.HasValue ? (bool?)(value.Value == 1) : null;
        }

        private static void ApplySectionCount(FamilyManager manager, int sectionCount)
        {
            if (sectionCount < 1 || sectionCount > 9)
                throw new ArgumentException("В исходном семействе доступны от 1 до 9 промежуточных секций.");
            SetIntegerValue(manager, "Секции_Промежуточные_Количество", sectionCount);
        }

        private static Action ApplyValveState(FamilyManager manager, bool enabled)
        {
            var parameter = manager.get_Parameter(ValveControl.ParameterName);
            if (parameter == null || parameter.StorageType != StorageType.Integer)
                throw new InvalidOperationException("В семействе нет логического параметра «" + ValveControl.ParameterName + "».");
            int requested = enabled ? 1 : 0;
            if (manager.CurrentType.AsInteger(parameter) == requested)
                return null;
            if (!parameter.IsDeterminedByFormula)
            {
                SetIntegerValue(manager, ValveControl.ParameterName, requested);
                return null;
            }

            bool? constant = ValveControl.ConstantFormula(parameter.Formula);
            if (!constant.HasValue || parameter.IsReadOnly || parameter.IsReporting)
                throw new InvalidOperationException("Нельзя переключить клапан: параметр задан зависимой формулой или доступен только для чтения.");
            var target = manager.CurrentType;
            var preserved = manager.Types.Cast<FamilyType>().Where(t => t.Name != target.Name)
                .Select(t => new { Type = t, Value = t.AsInteger(parameter) ?? (constant.Value ? 1 : 0) }).ToList();
            // Постоянная формула действует на всё семейство. Снимаем только её и сохраняем
            // прежние значения других типов, включая шаблон. Пересчёт — в общем Commit.
            manager.SetFormula(parameter, null);
            try
            {
                foreach (var item in preserved)
                {
                    if (item.Type.AsInteger(parameter) == item.Value)
                        continue;
                    manager.CurrentType = item.Type;
                    SetIntegerValue(manager, ValveControl.ParameterName, item.Value);
                }
            }
            finally { manager.CurrentType = target; }
            SetIntegerValue(manager, ValveControl.ParameterName, requested);
            return () =>
            {
                foreach (var item in preserved)
                    if (item.Type.AsInteger(parameter) != item.Value)
                        throw new InvalidOperationException("Изменилось состояние клапана у другого типа «" + item.Type.Name + "». Изменения отменены.");
            };
        }

        private static FamilyParameter FindFamilyParameter(FamilyManager manager, string name)
        {
            var matches = FamilyParameterNames.Aliases(name).Select(manager.get_Parameter).Where(p => p != null).ToList();
            if (matches.Count > 1)
                throw new InvalidOperationException("Неоднозначное соответствие параметра «" + name + "»: "
                + string.Join(", ", matches.Select(p => p.Definition.Name)));
            return matches.FirstOrDefault();
        }

        private static FamilyParameter LengthParameter(FamilyManager manager, string name)
        {
            var parameter = FindFamilyParameter(manager, name);
            if (parameter == null || parameter.StorageType != StorageType.Double)
                throw new InvalidOperationException("В семействе нет параметра размера «" + name + "».");
            return parameter;
        }

        private static double ReadLengthMm(FamilyManager manager, FamilyType type, string name)
        {
            double? value = type.AsDouble(LengthParameter(manager, name));
            if (!value.HasValue || !(FamilyParameterNames.IsOffset(name) ? DimensionValue.IsFinite(value.Value) : DimensionValue.IsPositive(value.Value)))
                throw new InvalidOperationException("В исходном типе не задан корректный размер «" + name + "».");
            return Common.IDHelper.ConvertInternalToMm(value.Value);
        }

        private static double ReadSourceLengthMm(FamilyManager manager, FamilyType type, string name)
        {
            double? value = type.AsDouble(LengthParameter(manager, name));
            // Незаданные значения отображаются пустыми. Подставных размеров здесь нет.
            return value.HasValue && DimensionValue.IsFinite(value.Value)
                ? Common.IDHelper.ConvertInternalToMm(value.Value) : double.NaN;
        }

        private static void SetLengthMm(FamilyManager manager, string name, double millimeters)
        {
            if (!(FamilyParameterNames.IsOffset(name) ? DimensionValue.IsFinite(millimeters) : DimensionValue.IsPositive(millimeters)))
                throw new ArgumentException("«" + name + "»: " + (FamilyParameterNames.IsOffset(name) ? "введите числовое смещение." : "размер должен быть больше нуля."));
            var parameter = LengthParameter(manager, name);
            double value = Common.IDHelper.ConvertMmToInternal(millimeters);
            double? current = manager.CurrentType.AsDouble(parameter);
            if (current.HasValue && Math.Abs(current.Value - value) < 1e-9)
                return;
            if (parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting)
                return;
            EnsureWritable(parameter);
            try
            {
                manager.Set(parameter, value);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Параметр «" + name + "», размер "
                + millimeters.ToString(CultureInfo.CurrentCulture) + " мм: " + OperationError.Message(ex), ex);
            }
        }

        private static void ApplyDimensions(FamilyManager manager, InstallationConfiguration configuration)
        {
            foreach (var dimension in configuration.DimensionAssignments())
                SetLengthMm(manager, dimension.Key, dimension.Value);
        }

        private static FamilyParameter InfoParameter(FamilyManager manager, string name)
        {
            var parameter = FindFamilyParameter(manager, name);
            if (parameter != null && parameter.StorageType != StorageType.String)
                throw new InvalidOperationException("Параметр «" + name + "» должен быть текстовым.");
            return parameter;
        }

        private static bool TryGetSystemNameInputs(FamilyManager manager, FamilyParameter output, out FamilyParameter value, out FamilyParameter mode)
        {
            value = mode = null;
            if (output == null || output.Definition.Name != InstallationInfo.SystemNameParameter || !output.IsDeterminedByFormula)
                return false;
            string expected = "if(" + InstallationInfo.SystemNameByTypeParameter + "," + InstallationInfo.SystemNameTypeParameter
                + "," + InstallationInfo.SystemNameInstanceParameter + ")";
            if (!string.Equals(Regex.Replace(output.Formula ?? "", @"\s+", ""), Regex.Replace(expected, @"\s+", ""), StringComparison.OrdinalIgnoreCase))
                return false;
            value = InfoParameter(manager, InstallationInfo.SystemNameTypeParameter);
            mode = FindFamilyParameter(manager, InstallationInfo.SystemNameByTypeParameter);
            if (value == null || mode == null || mode.StorageType != StorageType.Integer)
                throw new InvalidOperationException("Для поля «Имя системы» нужны текстовый параметр «" + InstallationInfo.SystemNameTypeParameter
                    + "» и логический параметр «" + InstallationInfo.SystemNameByTypeParameter + "». Проверьте основу семейства.");
            return true;
        }

        private static string ReadInfoValue(FamilyManager manager, FamilyType type, string name)
        {
            var parameter = InfoParameter(manager, name);
            FamilyParameter value, mode;
            if (TryGetSystemNameInputs(manager, parameter, out value, out mode))
                parameter = value;
            return parameter == null ? "" : type.AsString(parameter) ?? "";
        }

        private static void ApplyInfo(FamilyManager manager, InstallationInfo info)
        {
            foreach (var pair in info.Assignments())
            {
                var parameter = InfoParameter(manager, pair.Key);
                if (parameter == null)
                {
                    if (pair.Value.Length == 0)
                        continue;
                    throw new InvalidOperationException("В семействе отсутствует текстовый параметр «" + pair.Key
                        + "». Добавьте соответствующий параметр в исходное семейство, чтобы сохранить это поле.");
                }
                FamilyParameter value, mode;
                bool routed = TryGetSystemNameInputs(manager, parameter, out value, out mode);
                if (!routed)
                    value = parameter;
                bool writeValue = (manager.CurrentType.AsString(value) ?? "") != pair.Value;
                bool writeMode = routed && manager.CurrentType.AsInteger(mode) != 1;
                // Проверяем оба входа до записи. Формулу результата не снимаем;
                // значение «Имя системы_Экземпляр» и другие типы не изменяем.
                if (writeValue)
                    EnsureWritable(value);
                if (writeMode)
                    EnsureWritable(mode);
                try
                {
                    if (writeValue)
                        manager.Set(value, pair.Value);
                    if (writeMode)
                        manager.Set(mode, 1);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Поле «" + pair.Key + "»: " + OperationError.Message(ex), ex);
                }
            }
        }

        private static void VerifyInfoValues(FamilyManager manager, InstallationInfo info)
        {
            foreach (var pair in info.Assignments())
            {
                var output = InfoParameter(manager, pair.Key);
                FamilyParameter value, mode;
                if (TryGetSystemNameInputs(manager, output, out value, out mode))
                {
                    if (manager.CurrentType.AsInteger(mode) != 1)
                        ValueMismatch(mode, "1", manager.CurrentType.AsInteger(mode)?.ToString() ?? "не задано");
                    string input = manager.CurrentType.AsString(value) ?? "";
                    if (input != pair.Value)
                        ValueMismatch(value, pair.Value, input);
                }
                // После Commit сверяем именно результат Revit, а не только записанный вход.
                string actual = output == null ? "" : manager.CurrentType.AsString(output) ?? "";
                if (actual != pair.Value)
                {
                    if (output == null)
                        throw new InvalidOperationException("Не найден параметр «" + pair.Key + "».");
                    ValueMismatch(output, pair.Value, actual);
                }
            }
        }

        private static void RefreshDimensionModes(FamilyManager manager, InstallationConfiguration configuration)
        {
            configuration.RefreshDimensionModes(name =>
            {
                var parameter = LengthParameter(manager, name);
                return parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting;
            });
            configuration.RefreshBooleanModes(name =>
            {
                var parameter = FindFamilyParameter(manager, name);
                if (parameter == null || parameter.StorageType != StorageType.Integer)
                    throw new InvalidOperationException("В семействе нет логического параметра «" + name + "».");
                return parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting;
            });
        }

        private static void ReadCalculatedDimensions(Document document, InstallationConfiguration configuration)
        {
            RefreshDimensionModes(document.FamilyManager, configuration);
            foreach (var dimension in configuration.Dimensions())
                if (dimension.Value.IsCalculated)
                    dimension.Value.ReadResult(ReadLengthMm(document.FamilyManager, document.FamilyManager.CurrentType, dimension.Key));
            foreach (var flag in configuration.SharedBooleans)
                if (flag.Value.IsCalculated)
                    flag.Value.ReadResult(ReadBooleanValue(document.FamilyManager, document.FamilyManager.CurrentType, flag.Key));
        }

        private static string RequestedValuesReport(InstallationConfiguration configuration)
        {
            var lines = new List<string> { "Заданные значения:", "Секции_Промежуточные_Количество = " + configuration.IntermediateCount,
                    "Соединитель_Приточный_Клапан = " + (configuration.HasValve ? "Да" : "Нет") };
            lines.AddRange(configuration.BooleanAssignments().Select(p => p.Key + " = " + (p.Value != 0 ? "Да" : "Нет")));
            lines.AddRange(configuration.Info.Assignments().Select(p => p.Key + " = " + p.Value));
            lines.AddRange(configuration.SectionAssignments().Select(p => InstallationConfiguration.ParameterName(p.Key) + " = " + p.Value.DisplayName));
            lines.AddRange(configuration.DimensionAssignments().Select(p => (configuration.Catalog.Dependencies.Find(p.Key)?.Name ?? p.Key)
                + " = " + p.Value.ToString(CultureInfo.CurrentCulture) + " мм"));
            configuration.RecalculateLocal();
            lines.Add(configuration.Geometry.Report());
            return string.Join("\n", lines);
        }

        private static void ValueMismatch(FamilyParameter parameter, string requested, string actual)
        {
            throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "»: задано «" + requested
                + "», после пересчёта Revit получилось «" + actual + "»."
                + (parameter.IsDeterminedByFormula ? " Формула: " + parameter.Formula : " Проверьте зависимости этого параметра."));
        }

        private static void VerifyRequestedValues(Document document, InstallationConfiguration configuration)
        {
            var manager = document.FamilyManager;
            var type = manager.CurrentType;
            VerifyInfoValues(manager, configuration.Info);
            var integers = new Dictionary<string, int> { { "Секции_Промежуточные_Количество", configuration.IntermediateCount },
                    { "Соединитель_Приточный_Клапан", configuration.HasValve ? 1 : 0 } };
            foreach (var flag in configuration.BooleanAssignments())
                integers.Add(flag.Key, flag.Value);
            foreach (var pair in integers)
            {
                var parameter = manager.get_Parameter(pair.Key);
                int? actual = type.AsInteger(parameter);
                if (actual != pair.Value)
                    ValueMismatch(parameter, pair.Value.ToString(), actual.HasValue ? actual.Value.ToString() : "не задано");
            }
            foreach (var pair in configuration.DimensionAssignments())
            {
                var parameter = LengthParameter(manager, pair.Key);
                double? actual = type.AsDouble(parameter);
                if (parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting)
                    continue;
                double requested = Common.IDHelper.ConvertMmToInternal(pair.Value);
                if (!actual.HasValue || double.IsNaN(actual.Value) || Math.Abs(actual.Value - requested) > 1e-9)
                    ValueMismatch(parameter, pair.Value.ToString(CultureInfo.CurrentCulture) + " мм", actual.HasValue
                        ? Common.IDHelper.ConvertInternalToMm(actual.Value).ToString(CultureInfo.CurrentCulture) + " мм" : "не задано");
            }
            foreach (var pair in configuration.SectionAssignments())
            {
                var parameter = SectionParameter(manager, pair.Key);
                var id = type.AsElementId(parameter);
                var actual = id == null || id.Equals(ElementId.InvalidElementId) ? null : DescribeType(document, id);
                if (actual == null || actual.Key != pair.Value.Key)
                    ValueMismatch(parameter, pair.Value.DisplayName, actual?.DisplayName ?? "не задано");
            }
            // При сохранении сверяем перенос правил с Revit. Во время ввода API не вызывается.
            configuration.Geometry = LocalGeometry.Calculate(configuration);
            if (string.IsNullOrWhiteSpace(configuration.Geometry.Limitation))
                foreach (var pair in configuration.Geometry.EndsMm)
                {
                    var parameter = manager.get_Parameter("Системный_" + pair.Key + "_Блок_Длина");
                    if (parameter == null || !pair.Value.HasValue)
                        continue;
                    double? actual = type.AsDouble(parameter);
                    if (!actual.HasValue || Math.Abs(Common.IDHelper.ConvertInternalToMm(actual.Value) - pair.Value.Value) > 0.001)
                        ValueMismatch(parameter, pair.Value.Value.ToString(CultureInfo.CurrentCulture) + " мм (локальный расчёт)",
                            actual.HasValue ? Common.IDHelper.ConvertInternalToMm(actual.Value).ToString(CultureInfo.CurrentCulture) + " мм" : "не задано");
                }
        }
    }
}

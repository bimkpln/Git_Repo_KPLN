using Autodesk.Revit.DB;
using KPLN_MEPBender.Common;
using System;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Parameters
{
    public sealed class ParameterSnapshotService
    {
        private static readonly Dictionary<string, BuiltInParameter> SystemParametersToTransfer = new Dictionary<string, BuiltInParameter>
        {
            { "Комментарий", BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS },
            { "Диаметр трубы", BuiltInParameter.RBS_PIPE_DIAMETER_PARAM },
            { "Диаметр кривой MEP", BuiltInParameter.RBS_CURVE_DIAMETER_PARAM },
            { "Ширина кривой MEP", BuiltInParameter.RBS_CURVE_WIDTH_PARAM },
            { "Высота кривой MEP", BuiltInParameter.RBS_CURVE_HEIGHT_PARAM },
            { "Ширина кабельного лотка", BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM },
            { "Высота кабельного лотка", BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM }
        };

        private static readonly HashSet<string> ParameterNamesToIgnore = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "КП_Размер_Текст",
            "КП_О_Сортировка",
        };

        public ParameterSnapshot Capture(Element source)
        {
            ParameterSnapshot snapshot = new ParameterSnapshot();

            foreach (Parameter parameter in source.Parameters)
            {
                if (!CanCapture(parameter))
                    continue;

                if (IsIgnoredParameterName(parameter.Definition.Name))
                    continue;

                int parameterId = parameter.Id.GetStableIntegerValue();
                if (parameterId <= 0)
                    continue;

                AddSnapshotValue(snapshot, parameter, $"ParameterId:{parameterId}", parameter.Definition.Name, parameterId, null);
            }

            foreach (KeyValuePair<string, BuiltInParameter> pair in SystemParametersToTransfer)
            {
                Parameter parameter = source.get_Parameter(pair.Value);
                if (!CanCapture(parameter))
                    continue;

                if (IsIgnoredParameterName(parameter.Definition.Name) || IsIgnoredParameterName(pair.Key))
                    continue;

                AddSnapshotValue(snapshot, parameter, $"BuiltIn:{(int)pair.Value}", parameter.Definition.Name, null, pair.Value);
            }

            return snapshot;
        }

        public void Apply(Element target, ParameterSnapshot snapshot)
        {
            if (target == null || snapshot == null)
                return;

            foreach (ParameterSnapshotValue value in snapshot.Values.Values)
            {
                if (IsIgnoredParameterName(value.Name))
                    continue;

                Parameter targetParameter = FindTargetParameter(target, value);
                if (!CanApply(targetParameter, value.StorageType))
                    continue;

                if (IsIgnoredParameterName(targetParameter.Definition.Name))
                    continue;

                SetParameterValue(targetParameter, value);
            }
        }

        private bool CanCapture(Parameter parameter)
        {
            return parameter != null
                   && parameter.HasValue
                   && parameter.Definition != null
                   && parameter.StorageType != StorageType.None;
        }

        private bool CanApply(Parameter parameter, StorageType expectedStorageType)
        {
            return parameter != null
                   && !parameter.IsReadOnly
                   && parameter.Definition != null
                   && parameter.StorageType == expectedStorageType
                   && parameter.StorageType != StorageType.None;
        }

        private void AddSnapshotValue(
            ParameterSnapshot snapshot,
            Parameter parameter,
            string key,
            string name,
            int? parameterId,
            BuiltInParameter? builtInParameter)
        {
            if (snapshot.Values.ContainsKey(key))
                return;

            ParameterSnapshotValue value = new ParameterSnapshotValue
            {
                Key = key,
                Name = name,
                StorageType = parameter.StorageType,
                ParameterIdInteger = parameterId,
                BuiltInParameter = builtInParameter
            };

            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    value.DoubleValue = parameter.AsDouble();
                    break;
                case StorageType.ElementId:
                    value.ElementIdValue = parameter.AsElementId();
                    break;
                case StorageType.Integer:
                    value.IntegerValue = parameter.AsInteger();
                    break;
                case StorageType.String:
                    value.StringValue = parameter.AsString();
                    break;
            }

            snapshot.Values.Add(key, value);
        }

        private Parameter FindTargetParameter(Element target, ParameterSnapshotValue value)
        {
            if (value.BuiltInParameter.HasValue)
                return target.get_Parameter(value.BuiltInParameter.Value);

            if (value.ParameterIdInteger.HasValue)
            {
                foreach (Parameter parameter in target.Parameters)
                {
                    if (parameter?.Id.GetStableIntegerValue() == value.ParameterIdInteger.Value)
                        return parameter;
                }
            }

            if (!string.IsNullOrWhiteSpace(value.Name))
                return target.LookupParameter(value.Name);

            return null;
        }

        private void SetParameterValue(Parameter parameter, ParameterSnapshotValue value)
        {
            switch (value.StorageType)
            {
                case StorageType.Double:
                    parameter.Set(value.DoubleValue);
                    break;
                case StorageType.ElementId:
                    parameter.Set(value.ElementIdValue);
                    break;
                case StorageType.Integer:
                    parameter.Set(value.IntegerValue);
                    break;
                case StorageType.String:
                    parameter.Set(value.StringValue ?? string.Empty);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private bool IsIgnoredParameterName(string parameterName)
        {
            return !string.IsNullOrWhiteSpace(parameterName)
                   && ParameterNamesToIgnore.Contains(parameterName);
        }
    }
}

// Версионные ветки сосредоточены только в этом файле.
// Используйте символ RevitYYYY или DebugYYYY для версии подключённых RevitAPI.dll.
// Если символ не задан, сохраняется современная ветка (Revit 2024+).
#if Revit2020 || Debug2020 || Revit2021 || Debug2021 || Revit2022 || Debug2022 || Revit2023 || Debug2023
#define KPLN_ELEMENT_ID_INT32
#endif
#if Revit2020 || Debug2020 || Revit2021 || Debug2021 || Revit2022 || Debug2022
#define KPLN_FILTER_CASE_SENSITIVE
#endif
#if Revit2020 || Debug2020
#define KPLN_LEGACY_UNITS
#endif

using Autodesk.Revit.DB;

namespace KPLN_Tools_OVVK.Common
{
    internal static class IDHelper
    {
        internal static string ParameterDataType(FamilyParameter parameter)
        { return ParameterDataType(parameter.Definition); }

        internal static string ParameterDataType(Definition definition)
        {
#if Revit2020 || Debug2020 || Revit2021 || Debug2021
            return definition.ParameterType.ToString();
#else
            return definition.GetDataType().TypeId;
#endif
        }

        internal static bool HaveCompatibleParameterDataTypes(Parameter source, FamilyParameter target)
        {
            if (source.StorageType != target.StorageType) return false;
#if Revit2020 || Debug2020 || Revit2021 || Debug2021
            return source.Definition.ParameterType == target.Definition.ParameterType;
#else
            var sourceType = source.Definition.GetDataType();
            var targetType = target.Definition.GetDataType();
            // TypeId — строка с версией схемы. NameEquals сравнивает сам тип данных,
            // поэтому длина старого RFA совместима с длиной обновлённого RVT.
            if (sourceType.NameEquals(targetType)) return true;
            // Разные спецификации линейных размеров (например, длина и размер трубы)
            // хранят значение в футах. Никакого повторного перевода из миллиметров здесь нет.
            return source.StorageType == StorageType.Double
                && UnitUtils.IsMeasurableSpec(sourceType) && UnitUtils.IsMeasurableSpec(targetType)
                && UnitUtils.IsValidUnit(sourceType, UnitTypeId.Millimeters)
                && UnitUtils.IsValidUnit(targetType, UnitTypeId.Millimeters);
#endif
        }

        internal static bool IsLengthParameter(FamilyParameter parameter)
        {
#if Revit2020 || Debug2020 || Revit2021 || Debug2021
            return parameter.Definition.ParameterType == ParameterType.Length;
#else
            return parameter.Definition.GetDataType() == SpecTypeId.Length;
#endif
        }

        internal static long ElIdValue(ElementId id)
        {
#if KPLN_ELEMENT_ID_INT32
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

        // Для 64-битных ID лучше использовать ElIdValue. checked предотвращает
        // молчаливое усечение больших идентификаторов в существующих int-вызовах.
        internal static int ElIdInt(ElementId id) { return checked((int)ElIdValue(id)); }

        internal static ElementId CreateElementId(long value)
        {
#if KPLN_ELEMENT_ID_INT32
            return new ElementId(checked((int)value));
#else
            return new ElementId(value);
#endif
        }

        internal static FilterRule CreateContainsRule(ElementId parameterId, string value)
        {
#if KPLN_FILTER_CASE_SENSITIVE
            return ParameterFilterRuleFactory.CreateContainsRule(parameterId, value, false);
#else
            return ParameterFilterRuleFactory.CreateContainsRule(parameterId, value);
#endif
        }

        internal static double ConvertMmToInternal(int valueMm) { return ConvertMmToInternal((double)valueMm); }

        internal static double ConvertMmToInternal(double valueMm)
        {
#if KPLN_LEGACY_UNITS
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#endif
        }

        internal static double ConvertInternalToMm(double valueInternal)
        {
#if KPLN_LEGACY_UNITS
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#endif
        }

        internal static double ConvertInternalAreaToSquareMeters(double valueInternal)
        {
#if KPLN_LEGACY_UNITS
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_SQUARE_METERS);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.SquareMeters);
#endif
        }

        // Единицы числовых параметров берём из самого семейства, включая электрические величины.
        internal static double FromParameterUnits(FamilyParameter parameter, double value)
        {
#if KPLN_LEGACY_UNITS
            return UnitUtils.ConvertToInternalUnits(value, parameter.DisplayUnitType);
#else
            return UnitUtils.ConvertToInternalUnits(value, parameter.GetUnitTypeId());
#endif
        }

        internal static double ToParameterUnits(FamilyParameter parameter, double value)
        {
#if KPLN_LEGACY_UNITS
            return UnitUtils.ConvertFromInternalUnits(value, parameter.DisplayUnitType);
#else
            return UnitUtils.ConvertFromInternalUnits(value, parameter.GetUnitTypeId());
#endif
        }

        internal static string ParameterUnitLabel(FamilyParameter parameter)
        {
#if KPLN_LEGACY_UNITS
            return LabelUtils.GetLabelFor(parameter.DisplayUnitType);
#else
            return LabelUtils.GetLabelForUnit(parameter.GetUnitTypeId());
#endif
        }
    }
}
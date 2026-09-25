using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class DuctProtectionParameters
    {
        private const string FireProtectionName = "ТС_Огнезащита";
        private const string SmokeProtectionName = "ТС_Дымозащита";

        internal static bool AreAvailable(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctInsulations)
                .WhereElementIsElementType()
                .Any(type => GetYesNoParameter(type, FireProtectionName) != null)
                && new FilteredElementCollector(doc)
                    .OfClass(typeof(MechanicalSystemType))
                    .Any(type => GetYesNoParameter(type, SmokeProtectionName) != null);
        }

        internal static bool TryIsProtected(Element host, Parameter systemTypeParameter,
            out bool isProtected, out string error)
        {
            isProtected = false;
            error = null;

            Element systemType = systemTypeParameter != null
                && systemTypeParameter.StorageType == StorageType.ElementId
                ? host.Document.GetElement(systemTypeParameter.AsElementId())
                : null;
            Parameter smokeProtection = GetYesNoParameter(systemType, SmokeProtectionName);
            if (smokeProtection == null)
            {
                error = $"У типа системы отсутствует параметр Да/Нет '{SmokeProtectionName}'. Толщина НЕ записана";
                return false;
            }

            isProtected = smokeProtection.AsInteger() == 1;
            foreach (ElementId insulationId in InsulationLiningBase.GetInsulationIds(host.Document, host.Id))
            {
                if (!(host.Document.GetElement(insulationId) is DuctInsulation insulation))
                    continue;

                Element insulationType = host.Document.GetElement(insulation.GetTypeId());
                Parameter fireProtection = GetYesNoParameter(insulationType, FireProtectionName);
                if (fireProtection == null)
                {
                    error = $"У типа изоляции отсутствует параметр Да/Нет '{FireProtectionName}'. Толщина НЕ записана";
                    return false;
                }

                isProtected |= fireProtection.AsInteger() == 1;
            }

            return true;
        }

        private static Parameter GetYesNoParameter(Element element, string name)
        {
            return element?.GetParameters(name).FirstOrDefault(parameter =>
                parameter.StorageType == StorageType.Integer
#if Debug2020 || Revit2020
                && parameter.Definition.ParameterType == ParameterType.YesNo);
#else
                && parameter.Definition.GetDataType() == SpecTypeId.Boolean.YesNo);
#endif
        }
    }
}

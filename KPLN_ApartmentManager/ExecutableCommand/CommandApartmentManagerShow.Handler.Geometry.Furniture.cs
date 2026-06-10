using Autodesk.Revit.DB;
using KPLN_ApartmentManager.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private static bool IsFurnitureOrPlumbingCategory(Category category)
        {
            if (category == null)
                return false;

            BuiltInCategory bic;
            try
            {
                bic = (BuiltInCategory)IDHelper.ElIdInt(category.Id);
            }
            catch
            {
                return false;
            }

            return bic == BuiltInCategory.OST_Furniture ||
                   bic == BuiltInCategory.OST_PlumbingFixtures;
        }

        private static List<FamilyInstance> FindFurnitureAndPlumbingSubComponentsRecursive(Document doc, FamilyInstance rootInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (doc == null || rootInstance == null)
                return result;

            CollectFurnitureAndPlumbingSubComponentsRecursive(doc, rootInstance, result);

            return result;
        }

        private static void CollectFurnitureAndPlumbingSubComponentsRecursive(
            Document doc,
            FamilyInstance current,
            List<FamilyInstance> result)
        {
            if (doc == null || current == null)
                return;

            ICollection<ElementId> subIds = current.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                if (IsFurnitureOrPlumbingCategory(subFi.Category))
                {
                    result.Add(subFi);
                    continue;
                }

                CollectFurnitureAndPlumbingSubComponentsRecursive(doc, subFi, result);
            }
        }

        private static Level ResolvePlacementLevelForNestedInstance(Document doc, FamilyInstance nestedFi, FamilyInstance apartmentFi)
        {
            if (doc == null)
                return null;

            ElementId nestedLevelId = GetInstanceLevelId(nestedFi);
            if (nestedLevelId != ElementId.InvalidElementId)
            {
                Level nestedLevel = doc.GetElement(nestedLevelId) as Level;
                if (nestedLevel != null)
                    return nestedLevel;
            }

            ElementId apartmentLevelId = GetInstanceLevelId(apartmentFi);
            if (apartmentLevelId != ElementId.InvalidElementId)
            {
                Level apartmentLevel = doc.GetElement(apartmentLevelId) as Level;
                if (apartmentLevel != null)
                    return apartmentLevel;
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .FirstOrDefault();
        }

        private static double GetRotationAngleOnXY(FamilyInstance fi)
        {
            if (fi == null)
                return 0.0;

            Transform tr = fi.GetTransform();
            if (tr == null)
                return 0.0;

            XYZ basisX = tr.BasisX;
            if (basisX == null)
                return 0.0;

            return Math.Atan2(basisX.Y, basisX.X);
        }

        private void CopyFurnitureAndPlumbingFromApartmentUnderlay(Document doc, FamilyInstance apartmentFi, List<string> debugMessages, List<ElementId> createdElementIds = null)
        {
            if (doc == null || apartmentFi == null)
                return;

            List<FamilyInstance> nestedItems = FindFurnitureAndPlumbingSubComponentsRecursive(doc, apartmentFi);
            if (nestedItems == null || nestedItems.Count == 0)
                return;

            foreach (FamilyInstance nestedFi in nestedItems)
            {
                if (nestedFi == null)
                    continue;

                try
                {
                    FamilySymbol symbol = nestedFi.Symbol;
                    if (symbol == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + " не найден тип. " + BuildFurnitureDebugElementLabel(nestedFi));
                        continue;
                    }

                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                        doc.Regenerate();
                    }

                    Transform tr = nestedFi.GetTransform();
                    if (tr == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + " не найден Transform. " + BuildFurnitureDebugElementLabel(nestedFi));
                        continue;
                    }

                    XYZ insertPoint = tr.Origin;
                    if (insertPoint == null)
                        continue;

                    Level level = ResolvePlacementLevelForNestedInstance(doc, nestedFi, apartmentFi);
                    if (level == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не найден уровень для вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + ". " + BuildFurnitureDebugElementLabel(nestedFi));
                        continue;
                    }

                    FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;
                    FamilyInstance created = null;

                    switch (placementType)
                    {
                        case FamilyPlacementType.ViewBased:
                            break;

                        case FamilyPlacementType.OneLevelBased:
                        case FamilyPlacementType.OneLevelBasedHosted:
                        case FamilyPlacementType.WorkPlaneBased:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;

                        default:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;
                    }

                    if (created == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add(
                                "Не удалось создать экземпляр для вложенного элемента ID = " +
                                IDHelper.ElIdValue(nestedFi.Id) + ". " +
                                BuildFurnitureDebugElementLabel(nestedFi) +
                                ", placementType = " + placementType + ".");
                        continue;
                    }

                    if (createdElementIds != null)
                        createdElementIds.Add(created.Id);

                    double angle = GetRotationAngleOnXY(nestedFi);
                    if (Math.Abs(angle) > 1e-9)
                    {
                        Line axis = Line.CreateBound(insertPoint, insertPoint + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, created.Id, axis, angle);
                    }

                    ApplyFamilyInstanceFlipState(nestedFi, created, debugMessages);
                    ApplyFamilyInstanceAxisFlipCorrection(nestedFi, created, debugMessages);
                    CopyFurnitureDimensionParameters(nestedFi, created, debugMessages);
                    CopyFurnitureYesNoParameter(nestedFi, created, debugMessages, "Раскладной");
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add(
                            "Ошибка копирования вложенного элемента ID = " +
                            IDHelper.ElIdValue(nestedFi.Id) + ": " +
                            ex.GetType().Name + ": " + ex.Message +
                            ". " + BuildFurnitureDebugElementLabel(nestedFi));
                }
            }
        }

        private static string BuildFurnitureDebugElementLabel(FamilyInstance fi)
        {
            if (fi == null)
                return "<null>";

            string familyName = "<нет семейства>";
            string typeName = "<нет типа>";
            string placementType = "<нет>";

            try
            {
                if (fi.Symbol != null)
                {
                    typeName = !string.IsNullOrWhiteSpace(fi.Symbol.Name) ? fi.Symbol.Name : "<без имени типа>";

                    if (fi.Symbol.Family != null)
                    {
                        familyName = !string.IsNullOrWhiteSpace(fi.Symbol.Family.Name)
                            ? fi.Symbol.Family.Name
                            : "<без имени семейства>";

                        placementType = fi.Symbol.Family.FamilyPlacementType.ToString();
                    }
                }
            }
            catch
            {
            }

            return "ID = " + IDHelper.ElIdValue(fi.Id) +
                   ", family = '" + familyName + "'" +
                   ", type = '" + typeName + "'" +
                   ", category = '" + GetCategoryName(fi.Category) + "'" +
                   ", placementType = " + placementType;
        }

        private static string GetCategoryName(Category category)
        {
            if (category == null)
                return "<нет>";

            try
            {
                return !string.IsNullOrWhiteSpace(category.Name) ? category.Name : "<без имени>";
            }
            catch
            {
                return "<ошибка чтения>";
            }
        }

        private static void CopyFurnitureDimensionParameters(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            CopyFurnitureLengthParameter(source, target, debugMessages, "Глубина", "Depth", "КП_Глубина", "ADSK_Размер_Глубина");
            CopyFurnitureLengthParameter(source, target, debugMessages, "Ширина", "Width", "КП_Ширина", "ADSK_Размер_Ширина");
            CopyFurnitureLengthParameter(source, target, debugMessages, "Длина", "Length", "КП_Длина", "ADSK_Размер_Длина");
        }

        private static void CopyFurnitureYesNoParameter(FamilyInstance source, FamilyInstance target, List<string> debugMessages, params string[] parameterNames)
        {
            if (source == null || target == null || parameterNames == null || parameterNames.Length == 0)
                return;

            bool sourceValue;
            if (!TryGetYesNoParamFromElementOrType(source, out sourceValue, parameterNames))
                return;

            if (TrySetYesNoParamOnElementOrType(target, sourceValue, parameterNames))
                return;

            if (debugMessages != null)
            {
                debugMessages.Add(
                    "Не удалось перенести параметр '" + parameterNames[0] +
                    "' у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) +
                    " в созданный экземпляр ID = " + IDHelper.ElIdValue(target.Id) + ".");
            }
        }

        private static void CopyFurnitureLengthParameter(FamilyInstance source, FamilyInstance target, List<string> debugMessages, params string[] parameterNames)
        {
            if (source == null || target == null || parameterNames == null || parameterNames.Length == 0)
                return;

            double sourceValue;
            if (!TryGetLengthParamFromElementOrType(source, out sourceValue, parameterNames))
                return;

            if (TrySetLengthParamOnElementOrType(target, sourceValue, parameterNames))
                return;

            if (debugMessages != null)
            {
                debugMessages.Add(
                    "Не удалось перенести параметр '" + parameterNames[0] +
                    "' у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) +
                    " в созданный экземпляр ID = " + IDHelper.ElIdValue(target.Id) + ".");
            }
        }

        private static bool TrySetLengthParamOnElementOrType(Element element, double valueInternal, params string[] parameterNames)
        {
            if (element == null || parameterNames == null || parameterNames.Length == 0)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = element.LookupParameter(parameterName);
                if (TrySetLengthParameter(p, valueInternal))
                    return true;
            }

            Element typeElem = null;
            if (element.Document != null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = element.Document.GetElement(typeId);
            }

            if (typeElem == null)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = typeElem.LookupParameter(parameterName);
                if (TrySetLengthParameter(p, valueInternal))
                    return true;
            }

            return false;
        }

        private static bool TryGetYesNoParamFromElementOrType(Element element, out bool value, params string[] parameterNames)
        {
            value = false;

            if (element == null || parameterNames == null || parameterNames.Length == 0)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = element.LookupParameter(parameterName);
                if (TryGetYesNoParameterValue(p, out value))
                    return true;
            }

            Element typeElem = null;
            if (element.Document != null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = element.Document.GetElement(typeId);
            }

            if (typeElem == null)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = typeElem.LookupParameter(parameterName);
                if (TryGetYesNoParameterValue(p, out value))
                    return true;
            }

            return false;
        }

        private static bool TrySetYesNoParamOnElementOrType(Element element, bool value, params string[] parameterNames)
        {
            if (element == null || parameterNames == null || parameterNames.Length == 0)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = element.LookupParameter(parameterName);
                if (TrySetYesNoParameter(p, value))
                    return true;
            }

            Element typeElem = null;
            if (element.Document != null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = element.Document.GetElement(typeId);
            }

            if (typeElem == null)
                return false;

            foreach (string parameterName in parameterNames)
            {
                Parameter p = typeElem.LookupParameter(parameterName);
                if (TrySetYesNoParameter(p, value))
                    return true;
            }

            return false;
        }

        private static bool TrySetYesNoParameter(Parameter parameter, bool value)
        {
            if (parameter == null || parameter.IsReadOnly)
                return false;

            try
            {
                if (parameter.StorageType == StorageType.Integer)
                {
                    parameter.Set(value ? 1 : 0);
                    return true;
                }

                if (parameter.StorageType == StorageType.String)
                {
                    parameter.Set(value ? "1" : "0");
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static bool TrySetLengthParameter(Parameter parameter, double valueInternal)
        {
            if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.Double)
                return false;

            try
            {
                parameter.Set(valueInternal);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ApplyFamilyInstanceAxisFlipCorrection(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            if (source == null || target == null)
                return;

            ApplyHandAxisFlipCorrection(source, target, debugMessages);
            ApplyFacingAxisFlipCorrection(source, target, debugMessages);
        }

        private static void ApplyHandAxisFlipCorrection(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            XYZ sourceHand = GetFamilyInstanceHandDirection2D(source, null);
            XYZ targetHand = GetFamilyInstanceHandDirection2D(target, null);

            if (!ShouldFlipFurnitureAxis(targetHand, sourceHand))
                return;

            if (!CanFlipHand(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно исправить HandOrientation у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipHand.");

                return;
            }

            try
            {
                target.flipHand();
                TryRegenerateFurnitureDocument(target.Document);
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось исправить HandOrientation у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static void ApplyFacingAxisFlipCorrection(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            XYZ sourceFacing = GetFamilyInstanceFacingDirection2D(source, null);
            XYZ targetFacing = GetFamilyInstanceFacingDirection2D(target, null);

            if (!ShouldFlipFurnitureAxis(targetFacing, sourceFacing))
                return;

            if (!CanFlipFacing(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно исправить FacingOrientation у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipFacing.");

                return;
            }

            try
            {
                target.flipFacing();
                TryRegenerateFurnitureDocument(target.Document);
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось исправить FacingOrientation у мебели/сантехники ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static bool ShouldFlipFurnitureAxis(XYZ targetDirection, XYZ sourceDirection)
        {
            if (targetDirection == null || sourceDirection == null)
                return false;

            double dot = Dot2D(targetDirection, sourceDirection);
            double cross = Math.Abs(Cross2D(targetDirection, sourceDirection));

            return dot < -0.98 && cross < 0.05;
        }

        private static void TryRegenerateFurnitureDocument(Document doc)
        {
            if (doc == null)
                return;

            try
            {
                doc.Regenerate();
            }
            catch
            {
            }
        }

        private static void ApplyFamilyInstanceFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            if (source == null || target == null)
                return;

            ApplyHandFlipState(source, target, debugMessages);
            ApplyFacingFlipState(source, target, debugMessages);
        }

        private static void ApplyHandFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            bool sourceValue;
            bool targetValue;

            if (!TryGetHandFlipped(source, out sourceValue))
                return;

            if (!TryGetHandFlipped(target, out targetValue))
                return;

            if (sourceValue == targetValue)
                return;

            if (!CanFlipHand(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно повторить flipHand у вложенного элемента ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipHand.");
                return;
            }

            try
            {
                target.flipHand();
                TryRegenerateFurnitureDocument(target.Document);
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось применить flipHand к вложенному элементу ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static void ApplyFacingFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            bool sourceValue;
            bool targetValue;

            if (!TryGetFacingFlipped(source, out sourceValue))
                return;

            if (!TryGetFacingFlipped(target, out targetValue))
                return;

            if (sourceValue == targetValue)
                return;

            if (!CanFlipFacing(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно повторить flipFacing у вложенного элемента ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipFacing.");
                return;
            }

            try
            {
                target.flipFacing();
                TryRegenerateFurnitureDocument(target.Document);
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось применить flipFacing к вложенному элементу ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static bool TryGetHandFlipped(FamilyInstance fi, out bool value)
        {
            value = false;

            if (fi == null)
                return false;

            try
            {
                value = fi.HandFlipped;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetFacingFlipped(FamilyInstance fi, out bool value)
        {
            value = false;

            if (fi == null)
                return false;

            try
            {
                value = fi.FacingFlipped;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
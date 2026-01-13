using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_OpeningHoleManager.Forms.SelectionFilters
{
    internal sealed class MechEquipLinkSelectionFilter : ISelectionFilter
    {
        private readonly Document _doc;
        private Document _linkDoc = null;

        public MechEquipLinkSelectionFilter(Document doc)
        {
            _doc = doc;
        }

        public bool AllowElement(Element elem) => true;

        public bool AllowReference(Reference reference, XYZ position)
        {
            if (_doc.GetElement(reference) is RevitLinkInstance rli)
                _linkDoc = rli.GetLinkDocument();
            else
            {
                _linkDoc = null;
                return false;
            }

            if (_doc.Title.Contains("СЕТ"))
            {
                if (_linkDoc.GetElement(reference.LinkedElementId) is FamilyInstance fi)
                    return 
                        fi.Category != null
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                        && (fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment || (fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_GenericModel))
#else
                        && (fi.Category.BuiltInCategory == BuiltInCategory.OST_MechanicalEquipment || (fi.Category.BuiltInCategory == BuiltInCategory.OST_GenericModel))
#endif
                        && (fi.Symbol.FamilyName.StartsWith("501_ЗИ_Отвер") || (fi.Symbol.FamilyName.StartsWith("ASML_О_Отверстие") && fi.Symbol.FamilyName.Contains("В стене")));

            }
            else
            {
                if (_linkDoc.GetElement(reference.LinkedElementId) is FamilyInstance fi)
                    return
                        fi.Category != null
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                        && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment
#else
                        && fi.Category.BuiltInCategory == BuiltInCategory.OST_MechanicalEquipment
#endif
                        && fi.Symbol.FamilyName.StartsWith("501_ЗИ_Отвер");
            }

            return false;
        }
    }
}

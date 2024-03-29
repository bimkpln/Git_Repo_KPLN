using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System.Collections.Generic;



namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandChangeLevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            FormChageLevel formChageLevel = new FormChageLevel(doc);
            formChageLevel.ShowDialog();

            return Result.Succeeded;
        }

        public static BuiltInParameter[] GetParametersForMovingItems(Element element)
        {
            string category = element.Category.Name;

            BuiltInParameter[] parameters;
            BuiltInParameter baseLevel;
            BuiltInParameter baseOffset; 
            BuiltInParameter topLevel;
            BuiltInParameter topOffset;

            if (category == "OST_Walls")
            {
                baseLevel = BuiltInParameter.WALL_BASE_CONSTRAINT;
                baseOffset = BuiltInParameter.WALL_BASE_OFFSET;
                topLevel = BuiltInParameter.WALL_HEIGHT_TYPE;
                topOffset = BuiltInParameter.WALL_TOP_OFFSET;

            }
            else if (category == "OST_StructuralColumns")
            {
                baseLevel = BuiltInParameter.FAMILY_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM;
                topLevel = BuiltInParameter.FAMILY_TOP_LEVEL_PARAM;
                topOffset = BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM;

            }
            else if (category == "OST_Floors")
            {
                baseLevel = BuiltInParameter.LEVEL_PARAM;
                baseOffset = BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM;

            }
            else if (category == "OST_Ceilings")
            {
                baseLevel = BuiltInParameter.LEVEL_PARAM;
                baseOffset = BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM;

            }
            else if (category == "OST_Roofs")
            {
                baseLevel = BuiltInParameter.ROOF_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM;
            }
            else if (category == "OST_Windows" || category == "OST_Doors")
            {
                baseLevel = BuiltInParameter.FAMILY_LEVEL_PARAM;
                baseOffset = BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM;
            }
            else if (category == "OST_Stairs" || category == "OST_Ramps")
            {
                baseLevel = BuiltInParameter.STAIRS_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.STAIRS_BASE_OFFSET;
                topLevel = BuiltInParameter.STAIRS_TOP_LEVEL_PARAM;
                topOffset = BuiltInParameter.STAIRS_TOP_OFFSET;
            }
            else if (category == "OST_StairsRailing")
            {
                baseLevel = BuiltInParameter.STAIRS_RAILING_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.STAIRS_RAILING_HEIGHT_OFFSET;
            }
            else
            {
                baseLevel = BuiltInParameter.FAMILY_LEVEL_PARAM;
                baseOffset = BuiltInParameter.INSTANCE_ELEVATION_PARAM;
            }

            return parameters;
        }
    }
}

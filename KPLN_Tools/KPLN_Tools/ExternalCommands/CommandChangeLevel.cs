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
            BuiltInParameter[] parameters = null;

            if (element == null || element.Category == null)
            {
                return parameters;
            }

            BuiltInParameter baseLevel;
            BuiltInParameter baseOffset; 
            BuiltInParameter topLevel;
            BuiltInParameter topOffset;

            string category = element.Category.Name;

            if (category == "Стены")
            {
                baseLevel = BuiltInParameter.WALL_BASE_CONSTRAINT;
                baseOffset = BuiltInParameter.WALL_BASE_OFFSET;
                topLevel = BuiltInParameter.WALL_HEIGHT_TYPE;
                topOffset = BuiltInParameter.WALL_TOP_OFFSET;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset, topLevel, topOffset };

            }
            else if (category == "Несущие колонны")
            {
                baseLevel = BuiltInParameter.FAMILY_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM;
                topLevel = BuiltInParameter.FAMILY_TOP_LEVEL_PARAM;
                topOffset = BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset, topLevel, topOffset };

            }
            else if (category == "Перекрытия")
            {
                baseLevel = BuiltInParameter.LEVEL_PARAM;
                baseOffset = BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };

            }
            else if (category == "Потолки")
            {
                baseLevel = BuiltInParameter.LEVEL_PARAM;
                baseOffset = BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };

            }
            else if (category == "Кровли")
            {
                baseLevel = BuiltInParameter.ROOF_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };
            }
            else if (category == "Окна" || category == "Двери")
            {
                baseLevel = BuiltInParameter.FAMILY_LEVEL_PARAM;
                baseOffset = BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };
            }
            else if (category == "Лестницы" || category == "Пандусы")
            {
                baseLevel = BuiltInParameter.STAIRS_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.STAIRS_BASE_OFFSET;
                topLevel = BuiltInParameter.STAIRS_TOP_LEVEL_PARAM;
                topOffset = BuiltInParameter.STAIRS_TOP_OFFSET;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset, topLevel, topOffset };
            }
            else if (category == "Ограждения лестниц")
            {
                baseLevel = BuiltInParameter.STAIRS_RAILING_BASE_LEVEL_PARAM;
                baseOffset = BuiltInParameter.STAIRS_RAILING_HEIGHT_OFFSET;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };
            }
            else
            {
                baseLevel = BuiltInParameter.FAMILY_LEVEL_PARAM;
                baseOffset = BuiltInParameter.INSTANCE_ELEVATION_PARAM;

                parameters = new BuiltInParameter[] { baseLevel, baseOffset };
            }

            return parameters;
        }
        
    }
}

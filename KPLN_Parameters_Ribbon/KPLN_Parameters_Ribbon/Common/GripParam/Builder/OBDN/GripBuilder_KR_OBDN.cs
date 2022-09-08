﻿using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Common.Tools;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder.OBDN
{
    internal class GripBuilder_KR_OBDN : GripBuilder_KR
    {
        
        public GripBuilder_KR_OBDN(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        { 
        }
        
        public GripBuilder_KR_OBDN(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        protected override void FloorNumberUnderLevel(ref int counter)
        {
            foreach (Element elem in _elemsUnderLevel)
            {
                Level baseLevel = LevelTool.GetLevelOfElement(elem, Doc);
                if (baseLevel != null)
                {
                    string floorNumber = null; 

                    floorNumber = LevelTool.GetFloorNumberDecrementLevel(baseLevel, LevelNumberIndex, Doc, SplitLevelChar);

                    if (floorNumber == null)
                    {
                        Print($"Не найден уровень выше, для уровня {baseLevel.Name} " +
                            $"при обработке элемента: {elem.Name} c id: {elem.Id.IntegerValue}." +
                            "\nДля уровней необходимо заполнить параметр: На уровень выше, за исключением последнего этажа",
                            KPLN_Loader.Preferences.MessageType.Error);

                        continue;
                    }
                    Parameter floor = elem.LookupParameter(LevelParamName);
                    if (floor == null) continue;
                    floor.Set(floorNumber);
                    counter++;
                }
                else
                {
                    Print($"Не найден уровень у элемента с Id: {elem.Id}", KPLN_Loader.Preferences.MessageType.Error);
                }
            }
        }
    }
}

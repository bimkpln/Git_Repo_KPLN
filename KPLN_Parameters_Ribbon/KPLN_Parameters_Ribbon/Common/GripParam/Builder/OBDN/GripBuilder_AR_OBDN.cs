using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
using static KPLN_Loader.Output.Output;


namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder.OBDN
{
    internal class GripBuilder_AR_OBDN : GripBuilder_AR
    {

        public GripBuilder_AR_OBDN(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
        }

        public GripBuilder_AR_OBDN(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : base(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName, splitLevelChar)
        {
        }

        protected override void FloorNumberOnLevelByElement(Progress_Single pb)
        {
            foreach (Element elem in ElemsOnLevel)
            {
                try
                {
                    Level baseLevel = LevelTool.GetLevelOfElement(elem, Doc);
                    if (baseLevel != null)
                    {
                        string floorNumber = LevelTool.GetFloorNumberByLevel(baseLevel, LevelNumberIndex, SplitLevelChar);

                        double offsetFromLev = LevelTool.GetElementLevelGrip(elem, baseLevel);

                        if (offsetFromLev < 0)
                        {
                            Level nearestBeelowLevel = LevelTool.GetNearestBelowLevel(baseLevel, Doc);
                            string nearestBeelowLevelNumber = LevelTool.GetFloorNumberByLevel(nearestBeelowLevel, LevelNumberIndex, SplitLevelChar);
                            // ОБДН: АНАЛИЗ БЛИЖАЙШЕГО УРОВНЯ С УЧЕТОМ ОФФСЕТА СНИЗУ. ЕСЛИ ОН С ТАКИМ ЖЕ ИНДЕКСОМ - ДЕКРЕМЕНТА НЕ ПРОХОДИТ
                            if (nearestBeelowLevelNumber != floorNumber)
                            {
                                floorNumber = LevelTool.GetFloorNumberDecrementLevel(baseLevel, LevelNumberIndex, SplitLevelChar);
                            }
                        }

                        if (floorNumber == null) continue;
                        Parameter floor = elem.LookupParameter(LevelParamName);
                        if (floor == null) continue;
                        floor.Set(floorNumber);
                        pb.Increment();
                    }
                }
                catch (Exception e)
                {
                    PrintError(e, "Не удалось обработать элемент: " + elem.Id.IntegerValue + " " + elem.Name);
                }
            }
        }

    }
}

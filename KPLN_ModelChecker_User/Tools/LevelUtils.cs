/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#region Utils
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
#endregion

namespace RevitElementsElevation
{
    public static class LevelUtils
    {
        /// <summary>
        /// Попытаться найти уровень, к которому привязан элемент
        /// </summary>
        /// <param name="fi"></param>
        /// <returns></returns>
        public static Level GetBaseLevelofElement(FamilyInstance fi)
        {
            Document doc = fi.Document;
            Level baseLevel = null;

            try
            {
                //Это для обычных семейств на основе стены или уровня, у которых есть параметр "Уровень"
                baseLevel = doc.GetElement(fi.LevelId) as Level;
            }
            catch { }
            if (baseLevel != null) return baseLevel;



            //Для семейств "По рабочей плоскости", установленных на уровень
            try
            {
                Element hostElem = fi.Host;
                baseLevel = hostElem as Level;
            }
            catch { }
            if (baseLevel != null) return baseLevel;



            return null;
        }



        /// <summary>
        /// Получает смещение от уровня через параметр семейства.
        /// 
        /// </summary>
        /// <param name="fi"></param>
        /// <returns></returns>
        public static double GetOffsetFromLevel(FamilyInstance fi)
        {
            double elev = -999;
            Parameter offsetParam = fi.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
            if (offsetParam != null)
            {
                elev = offsetParam.AsDouble();
                if (elev != -999) return elev;
            }

            offsetParam = fi.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
            if (offsetParam != null)
            {
                elev = offsetParam.AsDouble();
                if (elev != -999) return elev;
            }


            if (elev == 0)
            {
                offsetParam = fi.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
                elev = offsetParam.AsDouble();
            }

            return elev;
        }



        /// <summary>
        /// Находит ближайший снизу уровень от точки.
        /// </summary>
        /// <param name="point"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static Level GetNearestLevel(XYZ point, Document doc, double projectPointElevation)
        {
            double pointZ = point.Z;
            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToList();

            Level finalLevel = null;

            foreach (Level lev in levels)
            {
                if (finalLevel == null)
                {
                    finalLevel = lev;
                    continue;
                }
                if (lev.Elevation < finalLevel.Elevation)
                {
                    finalLevel = lev;
                    continue;
                }
            }

            double offset = 10000;
            foreach (Level lev in levels)
            {
                double levHeigth = lev.Elevation + projectPointElevation;
                double testElev = pointZ - levHeigth;
                if (testElev < 0) continue;

                if (testElev < offset)
                {
                    finalLevel = lev;
                    offset = testElev;
                }
            }

            return finalLevel;
        }

    }
}

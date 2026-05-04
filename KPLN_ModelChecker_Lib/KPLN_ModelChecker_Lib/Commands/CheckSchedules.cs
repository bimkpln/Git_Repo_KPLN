using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckSchedules : AbstrCheck
    {
        private static readonly int _filterWarnCount = 5;
        private static readonly int _filterErrorCount = 7;
        
        /// <summary>
        /// Допуск при проверках сразу в футах.
        /// </summary>
#if Debug2020 || Revit2020
        private static readonly double _gluedTolerance = UnitUtils.ConvertToInternalUnits(1.0, DisplayUnitType.DUT_MILLIMETERS);
#else
        private static readonly double _gluedTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
#endif

        public CheckSchedules() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка спецификаций";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CommandCheckSchedules",
                    new Guid("46808d8d-6447-4502-983c-819fcc5d0b1f"),
                    new Guid("46808d8d-6447-4502-983c-819fcc5d0b1e"));
        }

        public override Element[] GetElemsToCheck() =>
            new FilteredElementCollector(CheckDocument)
                .OfClass(typeof(ViewSheet))
                .Where(el => el is ViewSheet vsh && !vsh.IsTemplate && !vsh.IsPlaceholder)
                .ToArray();

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            ElementFilter shiFilter = new ElementClassFilter(typeof(ScheduleSheetInstance));
            ElementFilter fiFilter = new ElementClassFilter(typeof(FamilyInstance));
            foreach (var elem in elemColl)
            {
                if (!(elem is ViewSheet vsh))
                    continue;


                // Список инстансов спек
                ScheduleSheetInstance[] schInsts = vsh
                    .GetDependentElements(shiFilter)?
                    .Select(id => CheckDocument.GetElement(id))?
                    .Cast<ScheduleSheetInstance>()
                    .Where(ssi => !ssi.IsTitleblockRevisionSchedule)
                    .ToArray();


                #region Проверка кол-ва фильтров в спеке
                List<CheckerEntity> schTooManyFilters = SchedulesTooManyFilters(CheckDocument, schInsts);
                if (schTooManyFilters != null && schTooManyFilters.Any())
                    _checkerEntitiesCollHeap.AddRange(schTooManyFilters);
                #endregion


                #region Проверка на близкое расположение разных спек - потенциальная ошибка в фильтра
                // Если лист содержит осн надпись разрешения на измы - в игнор
                FamilyInstance[] fInst_izm = vsh
                    .GetDependentElements(fiFilter)?
                    .Select(id => CheckDocument.GetElement(id))?
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.FamilyName.Contains("Разрешение на внесение изм"))
                    .ToArray();
                if (fInst_izm.Length == 0)
                {
                    List<CheckerEntity> schFrmEquals = SchedulesFramesEqual(CheckDocument, vsh, schInsts);
                    if (schFrmEquals != null && schFrmEquals.Any())
                        _checkerEntitiesCollHeap.AddRange(schFrmEquals);
                }
                #endregion
            }


            return CheckResultStatus.Succeeded;
        }

        private List<CheckerEntity> SchedulesTooManyFilters(Document doc, ScheduleSheetInstance[] schInsts)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (var schinst in schInsts)
            {
                Element elem = doc.GetElement(schinst.ScheduleId);
                if (elem != null && elem is ViewSchedule vsch) 
                {
                    ScheduleDefinition def = vsch.Definition;
                    int filterCount = def.GetFilterCount();

                    if (filterCount >= _filterErrorCount)
                    {
                        result.Add(new CheckerEntity(
                            vsch,
                            $"Слишком много фильтров",
                            $"В спецификации \"{vsch.Name}\" задано {filterCount} фильтров (порог ошибки: {_filterErrorCount})",
                            $"Запрещено использовать такое количество фильтров. Обратись за консультацией в BIM-отдел"));
                    }
                    else if (filterCount >= _filterWarnCount)
                    {
                        result.Add(new CheckerEntity(
                            vsch,
                            $"Много фильтров",
                            $"В спецификации \"{vsch.Name}\" задано {filterCount} фильтров (порог предупреждения: {_filterWarnCount})",
                            $"Стоит проверить, нельзя ли упростить логику фильтрации.")
                            .Set_Status(ErrorStatus.Warning));
                    }
                }
            }

            return result;
        }

        private static List<CheckerEntity> SchedulesFramesEqual(Document doc, ViewSheet vsh, ScheduleSheetInstance[] schInsts)
        {
            // Меньше 2х - игнор
            if (schInsts.Length < 2)
                return null;


            // Кэшируем боксы и флаги Title, чтобы не пересчитывать в двойном цикле
            BoundingBoxXYZ[] boxes = new BoundingBoxXYZ[schInsts.Length];
            bool[] titles = new bool[schInsts.Length];
            for (int k = 0; k < schInsts.Length; k++)
                boxes[k] = GetScheduleBoxOnSheet(doc, vsh, schInsts[k], out titles[k]);

            List<CheckerEntity> result = new List<CheckerEntity>();
            for (int i = 0; i < schInsts.Length; i++)
            {
                var boxA = boxes[i];
                if (boxA == null)
                    continue;

                bool hasTitleA = titles[i];

                for (int j = i + 1; j < schInsts.Length; j++)
                {
                    var boxB = boxes[j];
                    if (boxB == null)
                        continue;

                    bool hasTitleB = titles[j];

                    // Беру координаты спек
                    // ПО факту точка вставки для спек -+2,1 мм на границы вида (статично)
                    double offset = 0.007;

                    XYZ centerA = (boxA.Min + boxA.Max) * 0.5;
                    XYZ centerB = (boxB.Min + boxB.Max) * 0.5;

                    XYZ a_LU = MoveInside(new XYZ(boxA.Min.X, boxA.Max.Y, 0), centerA, offset, hasTitleA);
                    XYZ a_LD = MoveInside(new XYZ(boxA.Min.X, boxA.Min.Y, 0), centerA, offset, false);
                    XYZ a_RU = MoveInside(new XYZ(boxA.Max.X, boxA.Max.Y, 0), centerA, offset, hasTitleA);
                    XYZ a_RD = MoveInside(new XYZ(boxA.Max.X, boxA.Min.Y, 0), centerA, offset, false);

                    XYZ b_LU = MoveInside(new XYZ(boxB.Min.X, boxB.Max.Y, 0), centerB, offset, hasTitleB);
                    XYZ b_LD = MoveInside(new XYZ(boxB.Min.X, boxB.Min.Y, 0), centerB, offset, false);
                    XYZ b_RU = MoveInside(new XYZ(boxB.Max.X, boxB.Max.Y, 0), centerB, offset, hasTitleB);
                    XYZ b_RD = MoveInside(new XYZ(boxB.Max.X, boxB.Min.Y, 0), centerB, offset, false);


                    if (new Outline(a_LD, a_RU).Intersects(new Outline(b_LD, b_RU), _gluedTolerance))
                    {
                        // Проверка на склеивание
                        double LDLUDist = a_LD.DistanceTo(b_LU);
                        double LULDDist = a_LU.DistanceTo(b_LD);
                        double RDRUDist = a_RD.DistanceTo(b_RU);
                        double RURDDist = a_RU.DistanceTo(b_RD);
                        if(LDLUDist <= _gluedTolerance
                            || LULDDist <= _gluedTolerance
                            || RDRUDist <= _gluedTolerance
                            || RURDDist <= _gluedTolerance)
                            result.Add(new CheckerEntity(
                                new Element[] { schInsts[i], schInsts[j] },
                                $"Склеивание на листе \"{vsh.Title}\"",
                                $"Две разные спецификации размещены рядом - попытка \"склеивания\" данных",
                                $"Склеивать разные спецификации не рекомендуется, т.к. может привести к разным настройкам фильтрации/сортирвки. " +
                                    $"Если это разные спецификации, которые визуально не являются продолжением одна другой - можно отправить в допуск.")
                                .Set_Status(ErrorStatus.Warning));
                        else if (new Outline(a_LD, a_RU).Intersects(new Outline(b_LD, b_RU), 0))
                            result.Add(new CheckerEntity(
                                new Element[] { schInsts[i], schInsts[j] },
                                $"Пересечение на листе \"{vsh.Title}\"",
                                $"Две разные спецификации пересекаются",
                                $"При экспорте в pdf будет наложение."));
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Собираю всю инфу за одну проходку
        /// </summary>
        private static BoundingBoxXYZ GetScheduleBoxOnSheet(Document doc, ViewSheet vsh, ScheduleSheetInstance ssi, out bool hasTitle)
        {
            hasTitle = false;
            if (!(doc.GetElement(ssi.ScheduleId) is ViewSchedule vsch))
                return null;

            hasTitle = vsch.Definition.ShowTitle;
            return ssi.get_BoundingBox(vsh);
        }

        private static XYZ MoveInside(XYZ point, XYZ center, double offset, bool hasTitle)
        {
            double x = point.X + Math.Sign(center.X - point.X) * offset;
            
            double y = point.Y + Math.Sign(center.Y - point.Y) * offset;
            if (hasTitle)
                y = point.Y;

            return new XYZ(x, y, 0);
        }
    }
}

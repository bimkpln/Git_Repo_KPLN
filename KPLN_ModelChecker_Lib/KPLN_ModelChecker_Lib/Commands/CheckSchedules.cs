using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
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

        public override Element[] GetElemsToCheck()
        {
            // Обрабатываю пользовательскую выборку листов
            List<Element> result = new List<Element>();
            List<ElementId> selIds = CheckUIApp.ActiveUIDocument.Selection.GetElementIds().ToList();
            if (selIds.Count > 0)
            {
                foreach (ElementId selId in selIds)
                {
                    Element elem = CheckDocument.GetElement(selId);

#if Debug2020 || Revit2020
                    int catId = elem.Category.Id.IntegerValue;
                    if (catId.Equals((int)BuiltInCategory.OST_Sheets))
#else
                    if (elem.Category.BuiltInCategory == BuiltInCategory.OST_Sheets)
#endif
                    {
                        ViewSheet curViewSheet = elem as ViewSheet;
                        result.Add(curViewSheet);
                    }
                }

                if (result.Count == 0)
                {
                    TaskDialog.Show(
                        "Ошибка", 
                        "В выборке нет ни одного листа. Либо выбери листы, либо сними выборку, чтобы проанализировались вообще все листы", 
                        TaskDialogCommonButtons.Ok);

                    WarningIfNoElemsOnModel = false;
                    return new Element[0];
                }
                else
                    return result.ToArray();
            }
            
            
            return new FilteredElementCollector(CheckDocument)
                .OfClass(typeof(ViewSheet))
                .Where(el => el is ViewSheet vsh && !vsh.IsTemplate && !vsh.IsPlaceholder)
                .ToArray();
        }

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


                #region Проверка кол-ва и качества фильтров в спеке
                List<CheckerEntity> schTooManyFilters = SchedulesCheckFilters(CheckDocument, schInsts);
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

        private List<CheckerEntity> SchedulesCheckFilters(Document doc, ScheduleSheetInstance[] schInsts)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            // Метка проверки спек ИОС
            string fileFullName = SQLiteDocService.GetFileFullName(doc);
            var dbDoc = SQLiteMainService.SQLiteDocServiceInst.GetDBDocuments_ByFileFullPath(fileFullName);
            bool isIOS = dbDoc != null && dbDoc.SubDepartmentId != 2 && dbDoc.SubDepartmentId != 3 && dbDoc.SubDepartmentId != 8;

            // Проход по всем спекам на листе
            foreach (var schinst in schInsts)
            {
                Element elem = doc.GetElement(schinst.ScheduleId);
                if (elem != null && elem is ViewSchedule vsch) 
                {
                    string schNameLowerCase = vsch.Name.ToLower();
                    ScheduleDefinition def = vsch.Definition;
                    int filterCount = def.GetFilterCount();

                    // Проверяю кол-во фильтров в спеке
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

                    // Проверяю наличие нужных фильтров ИОС в спеке
                    if (isIOS)
                    {
                        string paramName = "КП_И_Включение в спецификацию";
                        string paramUserDescr = $"\"{paramName} -> Меньше или равно -> Да\"";

                        // Флаг под проекты
                        bool isSET = doc.Title.StartsWith("СЕТ_1");
                        if (isSET)
                        {
                            paramName = "СМ_Смета";
                            paramUserDescr = $"\"{paramName} -> не равно -> Нет\" ИЛИ \"{paramName} -> Меньше или равно -> Да\"";
                        }

                        bool haveFilterWithIOSField = false;
                        var schFilters = def.GetFilters();
                        for (int i = 0; i < schFilters.Count; i++)
                        {
                            var schFilterId = schFilters[i].FieldId;
                            var schField = def.GetField(schFilterId);
                            if (schField == null)
                                continue;

                            if (doc.GetElement(schField.ParameterId) is SharedParameterElement shParam)
                            {
                                bool checkName = shParam.Name.Equals(paramName);
                                
                                bool checkGilterType = schFilters[i].FilterType == ScheduleFilterType.LessThanOrEqual;
                                bool checkFilterValue = schFilters[i].GetIntegerValue() == 1;
                                if (isSET && !checkGilterType && !checkFilterValue)
                                {
                                    checkGilterType = schFilters[i].FilterType == ScheduleFilterType.NotEqual;
                                    checkFilterValue = schFilters[i].GetIntegerValue() == 0;
                                }

                                haveFilterWithIOSField = checkName && checkGilterType && checkFilterValue;
                                
                                break;
                            }
                        }
                        
                        if(!haveFilterWithIOSField 
                            && !schNameLowerCase.Contains("эксплик")
                            && !schNameLowerCase.Contains("шапка спец")
                            && !schNameLowerCase.Contains("показатели систем")
                            && !schNameLowerCase.Contains("количество листов")
                            && !schNameLowerCase.Contains("документ"))
                        {
                            result.Add(new CheckerEntity(
                                vsch,
                                $"ИОС: Нет обязательного фильтра",
                                $"В спецификации \"{vsch.Name}\" нет фильтра по параметру \"{paramName}\", либо он добавлен с ошибкой в условиях",
                                $"Данный параметр снимает технические элементы и обязателен для всех спецификаций по объёмам для ИОС. " +
                                    $"Единственно правильный вариант, который ИСКЛЮЧАЕТ все технические элементы из спецификации это {paramUserDescr}")
                                .Set_Status(ErrorStatus.Warning));
                        }
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
                                $"Склеивать разные спецификации не рекомендуется, т.к. может привести к разным настройкам фильтрации/сортировки. " +
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

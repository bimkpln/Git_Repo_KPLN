using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_Lib.Services;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckHolesOnLists : AbstrCheck
    {
        /// <summary>
        /// Список имен ХОСТОВ для файлов АР, которые обрабатываются по правилу "начинается с"
        /// </summary>
        private readonly string[] _krHostNames_StartWith = new string[]
            {
                // Проекты КПЛН
                "00_",
                // Проекты СМЛТ
                "КЖ_",
            };

        /// <summary>
        /// Функция фильтрации отверстий в модели
        /// </summary>
        private Func<FamilyInstance, bool> _isHolesFunc;

        public CheckHolesOnLists() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка видимости отверстий";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckHoles",
                    new Guid("7db863cc-9856-4d22-9918-cfa38b29c05e"),
                    new Guid("7db863cc-9856-4d22-9918-cfa38b29c06e"));
        }

        public override Element[] GetElemsToCheck()
        {
            if (CheckDocument.Title.Contains("СЕТ_1")) 
            {
                _isHolesFunc = fi =>
                    fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows
                    && (fi.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие") 
                    || fi.Symbol.FamilyName.StartsWith("231_Отверстие") 
                    || fi.Symbol.FamilyName.StartsWith("231_Проем"));
            }
            else
            {
                _isHolesFunc = fi =>
                    fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment
                    && (fi.Symbol.FamilyName.StartsWith("199_Отверстие")
                    || fi.Symbol.FamilyName.StartsWith("231_Отверстие")
                    || fi.Symbol.FamilyName.StartsWith("231_Проем"));
            }

            FamilyInstance[] holesFamInsts = new FilteredElementCollector(CheckDocument)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(_isHolesFunc)
                .ToArray();

             return holesFamInsts;
        }

        /// <summary>
        /// Создать BBox по подрезкам вида
        /// </summary>
        private static BoundingBoxXYZ CreateViewBBox(View activeView)
        {
            // Если активна подрезка на плане - берем её (учитываем уровень с расширением)
            if (activeView is ViewPlan && activeView.CropBoxActive)
            {
                Level viewLevel = activeView.GenLevel;
                double levelZCoord = viewLevel.Elevation;
                BoundingBoxXYZ vCropBB = activeView.CropBox;
                XYZ vCropBBMin = vCropBB.Min;
                XYZ vCropBBMax = vCropBB.Max;

                return new BoundingBoxXYZ()
                {
                    Min = new XYZ(vCropBBMin.X, vCropBBMin.Y, levelZCoord - 10),
                    Max = new XYZ(vCropBBMax.X, vCropBBMax.Y, levelZCoord + 30),
                };
            }
            // Если активна подрезка на 3d - берем её
            else if (activeView is View3D view3D && view3D.IsSectionBoxActive)
            {
                BoundingBoxXYZ sectBox = view3D.GetSectionBox();
                Transform viewTrans = sectBox.Transform;
                return new BoundingBoxXYZ()
                {
                    Min = viewTrans.OfPoint(sectBox.Min),
                    Max = viewTrans.OfPoint(sectBox.Max),
                };
            }
            // Если активна подрезка на разрезе - берем её
            else if (activeView is ViewSection viewSection && activeView.CropBoxActive)
            {
                // В координатах плана
                var cropBox = viewSection.CropBox;
                double sectionDepth = Math.Abs(cropBox.Max.Z - cropBox.Min.Z);

                // В координатах самого окна
                var cropManager = viewSection.GetCropRegionShapeManager();
                IList<CurveLoop> cropShape = cropManager.GetCropShape();
                // Фрагментные разрезы сложно анализировать. Лучше пусть уходит на открытие
                if (cropShape.Count > 1)
                    return null;
                else
                {
                    XYZ viewDirection = viewSection.ViewDirection * sectionDepth;
                    cropShape.Add(CurveLoop.CreateViaTransform(cropShape.FirstOrDefault(), Transform.CreateTranslation(viewDirection.Negate())));
                }

                Solid fullCropSolid = GeometryCreationUtilities.CreateLoftGeometry(cropShape, new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId));

                return fullCropSolid.GetBoundingBox();
            }

            return null;
        }

        /// <summary>
        /// Получить уже промаркированный элемент ТИПОВОГО этажа (проверяю по экв. координатам по XY, и объёму солида отверстия)
        /// </summary>
        private static InstanceGeomData GetTypeFloorEqualElem(InstanceGeomData elemToCheckIGD, InstanceGeomData[] visibleIGDColl)
        {
            foreach (InstanceGeomData vIGD in visibleIGDColl)
            {
                XYZ elemToCheckCoordToVisibleElem = new XYZ(elemToCheckIGD.IGDGeomCenter.X, elemToCheckIGD.IGDGeomCenter.Y, vIGD.IGDGeomCenter.Z);

                if (elemToCheckCoordToVisibleElem.DistanceTo(vIGD.IGDGeomCenter) < 0.1
                    && Math.Abs(vIGD.IGDSolid.Volume - elemToCheckIGD.IGDSolid.Volume) < 0.1)
                    return vIGD;
            }

            return null;
        }

        /// <summary>
        /// Проверка наличия марки у элемента
        /// </summary>
        private static bool IsElemTaggedOnView(Document doc, Element elem, View view)
        {
            ElementFilter filter = new ElementClassFilter(typeof(IndependentTag));

            // Анализ на наличие марок у отверстий
            IList<ElementId> depElemIds = elem.GetDependentElements(filter);
            foreach (ElementId elId in depElemIds)
            {
                IndependentTag indTag = (IndependentTag)doc.GetElement(elId)
                    ?? throw new CheckerException($"Обратись к разработчику - не удалось получить марку. Id: {elId}");

                View tagView = (View)doc.GetElement(indTag.OwnerViewId)
                    ?? throw new CheckerException($"Обратись к разработчику - не удалось получить вид. Id: {elId}");

                if (tagView.Id == view.Id)
                    return true;
            }


            return false;
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            Selection selection = CheckUIApp.ActiveUIDocument.Selection;

            // Список видов, которые есть на листах
            ViewSheet[] selVSheets = selection
                .GetElementIds()
                .Select(id => CheckDocument.GetElement(id))
                .OfType<ViewSheet>()
                .ToArray();

            if (!selVSheets.Any())
            {
                TaskDialog taskDialog = new TaskDialog("Внимание!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Перед запуском нужно выбрать листы, которые будут анализироваться. " +
                        "\nВАЖНО, чтобы были выбраны ВСЕ листы, на которых отображаются отверстия, иначе могут быть ошибки в результате работы плагина",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                TaskDialogResult dialogResult = taskDialog.Show();

                return CheckResultStatus.Failed;
            }


            List<View> docViewColl = new List<View>();
            foreach(ViewSheet vSheet in selVSheets)
            {
                ICollection<ElementId> allViewPorts = vSheet.GetAllViewports();
                foreach (ElementId vpId in allViewPorts)
                {
                    Viewport vp = (Viewport)CheckDocument.GetElement(vpId);
                    View view = CheckDocument.GetElement(vp.ViewId) as View;
                    if (view.ViewType == ViewType.FloorPlan 
                        || view.ViewType == ViewType.Section 
                        || view.ViewType == ViewType.ProjectBrowser 
                        || view.ViewType == ViewType.Detail 
                        || view.ViewType == ViewType.EngineeringPlan)
                        docViewColl.Add(view);
                }
            }

            bool isKRDoc = false;
            string fileFullName = GetFileFullName(CheckDocument);
            DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(fileFullName);
            if (prjDBSubDepartment != null)
                isKRDoc = prjDBSubDepartment.Code == "КР";

            //_checkerEntitiesCollHeap.AddRange(CheckTaggedElems(elemColl, docViewColl, isKRDoc));
            _checkerEntitiesCollHeap.AddRange(CheckElemsVisibility(elemColl, docViewColl, isKRDoc));

            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Получить полное имя открытого файла
        /// </summary>
        /// <param name="doc">Документ Ревит для анализа</param>
        /// <returns></returns>
        private string GetFileFullName(Document doc) =>
            doc.IsWorkshared && !doc.IsDetached
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

        /// <summary>
        /// Получить список ошибок по отверстиям, которые не видны ни на одном виде
        /// </summary>
        private IEnumerable<CheckerEntity> CheckElemsVisibility(Element[] elemColl, IEnumerable<View> docViewColl, bool isKRModel)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            HashSet<ElementId> visibleIds = new HashSet<ElementId>();
            ElementFilter filter = new ElementClassFilter(typeof(Viewport));
            foreach (var view in docViewColl)
            {
                // Делаю проверку по подрезке вида, чтобы чекнуть, входит ли эл-нт в границы вида
                BoundingBoxXYZ viewBBox = CreateViewBBox(view);
                if (viewBBox != null)
                {
                    Outline viewExpandedOutline = GeometryWorker.CreateOutline_ByBBoxANDExpand(viewBBox, new XYZ(3, 3, 5));
                    BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(viewExpandedOutline, 0.1);
                    BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(viewExpandedOutline, 0.1);
                    LogicalOrFilter resFilter = new LogicalOrFilter(intersectsFilter, insideFilter);

                    FamilyInstance[] holesFamInstsOnViewCropColl = new FilteredElementCollector(CheckDocument)
                        .OfClass(typeof(FamilyInstance))
                        .WherePasses(resFilter)
                        .Cast<FamilyInstance>()
                        .Where(_isHolesFunc)
                        .ToArray();

                    // Если по фильтрам ничего не попало, то вид открывать не нужно
                    if (holesFamInstsOnViewCropColl == null || holesFamInstsOnViewCropColl.Length == 0)
                        continue;
                }

                FamilyInstance[] holesFamInstsOnViewColl = new FilteredElementCollector(CheckDocument, view.Id)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(_isHolesFunc)
                    .ToArray();


                foreach (FamilyInstance fi in holesFamInstsOnViewColl)
                {
                    if (IsElemTaggedOnView(CheckDocument, fi, view))
                        visibleIds.Add(fi.Id);
                }
            }

            InstanceGeomData[] instanceGeomDatas = null;
            foreach (Element elem in elemColl)
            {
                if (!(elem is FamilyInstance fi))
                    throw new CheckerException($"Обратись к разработчику - в анализ попало НЕ пользовательское семейство. Id: {elem.Id}");

                Element hostElem = fi.Host
                    ?? throw new CheckerException($"Обратись к разработчику - у отверстия нет основания. Дальнейший анализ невозможен. Id: {elem.Id}");

                bool isElemVisible = visibleIds.Count(id => id == elem.Id) > 0;
                if (isElemVisible)
                    continue;

                if (isKRModel)
                {
                    InstanceGeomData elIGD = new InstanceGeomData(elem);
                    if (instanceGeomDatas == null)
                        instanceGeomDatas = visibleIds
                            .Select(id => new InstanceGeomData(CheckDocument.GetElement(id)))
                            .ToArray();

                    // Если есть похожий элемент, то считаем, что всё ок.
                    InstanceGeomData typeFloorElem = GetTypeFloorEqualElem(elIGD, instanceGeomDatas);
                    if (typeFloorElem != null)
                        continue;
                }
                
                
                if (isKRModel || !_krHostNames_StartWith.Any(prefix => (bool)hostElem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix)))
                {
                    Level fiLvl = CheckDocument.GetElement(fi.LevelId) as Level;
                    
                    string info = $"Уровень элемента: \"{fiLvl.Name}\". Сделай так, чтобы отверстие было видно на необходимом листе." ;
                    if (!isKRModel)
                        info = $"Уровень элемента: \"{fiLvl.Name}\". Сделай так, чтобы отверстие было видно на необходимом листе." +
                            $"\nИНФО: Обрати внимание, что основа у элемента - \"{hostElem.Name}\".";

                    result.Add(new CheckerEntity(
                        elem,
                        "Отверстия нет на планах на листах модели",
                        $"Каждое отверстие должно хотя бы 1 раз встречаться на листе в модели и обязатольно иметь марку на таком виде",
                        info,
                        true));
                }
            }


            return result.OrderBy(ce => ((Level)CheckDocument.GetElement(ce.Element.LevelId)).Elevation);
        }

        /// <summary>
        /// Получить список ошибок по отверстиям, у которых НЕТ марки (это ошибка сама по себе, плюс - дорого искать вид на листе)
        /// </summary>
        private IEnumerable<CheckerEntity> CheckTaggedElems(Element[] elemColl, View[] docViewColl, bool isKRModel)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            ElementId[] docViewIdColl = docViewColl.Select(v => v.Id).ToArray();
            ElementFilter filter = new ElementClassFilter(typeof(IndependentTag));
            foreach (Element elem in elemColl)
            {
                if (!(elem is FamilyInstance fi))
                    throw new CheckerException($"Обратись к разработчику - в анализ попало НЕ пользовательское семейство. Id: {elem.Id}");

                Element hostElem = fi.Host
                    ?? throw new CheckerException($"Обратись к разработчику - у отверстия нет основания. Дальнейший анализ невозможен. Id: {elem.Id}");

                // Анализ на наличие марок у отверстий
                IList<ElementId> depElemIds = elem.GetDependentElements(filter);
                if (depElemIds.Count == 0)
                {
                    if (isKRModel 
                        || !_krHostNames_StartWith.Any(prefix => (bool)hostElem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix)))
                        result.Add(new CheckerEntity(
                            elem,
                            "Отверстие не промаркировано",
                            $"Все отверстия должны быть промаркированы, иначе нет связи между спецификацией и графикой",
                            "Найди подходящий вид, который обязательно размещён на листе, и добавь марку для отверстия" +
                                "\nВАЖНО: АР это, или КЖ - зависит от физической привязки к стене. Частая ошибка - отверстие размещают в слой отделки, а не в основную стену",
                            true));
                    else
                        result.Add(new CheckerEntity(
                            elem,
                            "Отверстие в КЖ не промаркировано",
                            $"Скорее всего, отверстия должны быть промаркированы и для КЖ. Если такой необходимости нет - группу можно игнорировать",
                            "Найди подходящий вид, который обязательно размещён на листе, и добавь марку для отверстия" +
                                "\nВАЖНО: АР это, или КЖ - зависит от физической привязки к стене. Частая ошибка - отверстие размещают в слой отделки, а не в основную стену",
                            true)
                            .Set_Status(ErrorStatus.Warning));

                    continue;
                }


                // Анализ марок на привязку к видам на листах
                bool tagNotOnList = true;
                foreach(ElementId elId in depElemIds)
                {
                    IndependentTag indTag = (IndependentTag)CheckDocument.GetElement(elId) 
                        ?? throw new CheckerException($"Обратись к разработчику - не удалось получить марку. Id: {elId}");
                    
                    View tagView = (View)CheckDocument.GetElement(indTag.OwnerViewId) 
                        ?? throw new CheckerException($"Обратись к разработчику - не удалось получить вид. Id: {elId}");

                    if (docViewIdColl.Contains(tagView.Id))
                        tagNotOnList = false;
                }

                if (tagNotOnList)
                {
                    if (isKRModel)
                        result.Add(new CheckerEntity(
                            elem,
                            "Марка отверстия НЕ размещена на лист",
                            $"Все отверстия должны быть промаркированы и размещены на лист, иначе нет связи между спецификацией и графикой",
                            "Найди подходящий вид и размести его на лист, либо промаркируй отверстие на уже сущестующем листе",
                            true));
                    else
                    {
                        if (_krHostNames_StartWith.Any(prefix => (bool)hostElem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix)))
                            result.Add(new CheckerEntity(
                                elem,
                                "Марка отверстия КЖ НЕ размещена на лист",
                                $"Скорее всего, отверстия должны быть и промаркированы, и вынесены на лист и для КЖ. Если такой необходимости нет - группу можно игнорировать",
                                "Найди подходящий вид и размести его на лист, либо промаркируй отверстие на уже сущестующем листе. " +
                                    "\nВАЖНО: АР это, или КЖ - зависит от физической привязки к стене. Частая ошибка - отверстие размещают в слой отделки, а не в основную стену",
                                true)
                                .Set_Status(ErrorStatus.Warning));
                        else
                            result.Add(new CheckerEntity(
                            elem,
                            "Марка отверстия АР НЕ размещена на лист",
                            $"Все отверстия должны быть промаркированы и размещены на лист, иначе нет связи между спецификацией и графикой",
                            "Найди подходящий вид и размести его на лист, либо промаркируй отверстие на уже сущестующем листе" +
                                "\nВАЖНО: АР это, или КЖ - зависит от физической привязки к стене. Частая ошибка - отверстие размещают в слой отделки, а не в основную стену",
                            true));
                    }
                }
            }

            return result;
        }
    }
}

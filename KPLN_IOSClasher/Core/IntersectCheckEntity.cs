using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность области анализа на коллизии
    /// </summary>
    internal class IntersectCheckEntity
    {
        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        private static readonly List<BuiltInCategory> _builtInCategories = new List<BuiltInCategory>()
        { 
            // ОВВК (ЭОМСС - огнезащита)
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_PipeCurves,
            // ЭОМСС
            BuiltInCategory.OST_CableTray,
        };

        private static LogicalOrFilter _elemCatLogicalOrFilter;
        private static Func<Element, bool> _elemFilterFunc;

        /// <summary>
        /// Фильтр для ключевой фильтрации по категориям
        /// </summary>
        public static LogicalOrFilter ElemCatLogicalOrFilter 
        { 
            get
            {
                if (_elemCatLogicalOrFilter == null)
                {
                    List<ElementFilter> catFilters = new List<ElementFilter>();
                    catFilters.AddRange(_builtInCategories.Select(bic => new ElementCategoryFilter(bic)));

                    _elemCatLogicalOrFilter = new LogicalOrFilter(catFilters);
                }

                return _elemCatLogicalOrFilter;
            } 
        }
        /// <summary>
        /// Общая функция для фильтра для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        public static Func<Element, bool> ElemExtraFilterFunc
        {
            get
            {
                if (_elemFilterFunc == null) 
                {
                    _elemFilterFunc = (el) =>
                        el.Category != null
                        && !(el is ElementType)
                        && !(el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("ASML_ОГК_")
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("Огнезащитный короб_EI150"));
                }

                return _elemFilterFunc;
            }
        }

        public IntersectCheckEntity(Document checkDoc, BoundingBoxXYZ filterBBox, Outline filterOutline)
        {
            CheckDoc = checkDoc;
            CheckBBox = filterBBox;
            CheckOutline = filterOutline;

            CheckDocTransform = CheckDoc.ActiveProjectLocation.GetTotalTransform();
            CheckDocBasePntPosition = new FilteredElementCollector(CheckDoc)
                .OfClass(typeof(BasePoint))
                .Cast<BasePoint>()
                .FirstOrDefault()
                .Position;

            SetCurrentDocElemsToCheck(CheckDoc);
        }

        public IntersectCheckEntity(Document checkDoc, BoundingBoxXYZ filterBBox, Outline filterOutline, RevitLinkInstance linkInst)
        {
            CheckDoc = checkDoc;
            CheckBBox = filterBBox;
            CheckOutline = filterOutline;

            CheckDocTransform = CheckDoc.ActiveProjectLocation.GetTotalTransform();
            CheckDocBasePntPosition = new FilteredElementCollector(CheckDoc)
                .OfClass(typeof(BasePoint))
                .Cast<BasePoint>()
                .FirstOrDefault()
                .Position;

            CheckLinkInst = linkInst;

            Document linkDoc = CheckLinkInst.GetLinkDocument();

            // Если открыто сразу несколько моделей одного проекта, то линки могут прилететь с другого файла. В таком случае - игнор и аннулирование CheckLinkInst
            if (linkDoc != null)
            {
                LinkBasePntPosition = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(BasePoint))
                    .Cast<BasePoint>()
                    .FirstOrDefault()
                    .Position;

                // Ищу результирующий Transform. Уточняю его при смещении БТП с нарушениями
                if (Math.Abs(CheckDocBasePntPosition.DistanceTo(LinkBasePntPosition)) > 0.1
                    && Math.Abs(CheckDocBasePntPosition.DistanceTo(new XYZ(0, 0, 0))) > 0.1)
                {
                    XYZ resultVect = CheckDocBasePntPosition - LinkBasePntPosition;
                    Transform difTransform = Transform.CreateTranslation(resultVect);
                    LinkTransfrom = difTransform;
                }
                else if (Math.Abs(LinkBasePntPosition.DistanceTo(new XYZ(0, 0, 0))) < 0.1)
                    LinkTransfrom = linkDoc.ActiveProjectLocation.GetTransform();
                else
                    LinkTransfrom = CheckDocTransform.Inverse * linkDoc.ActiveProjectLocation.GetTransform();

                SetCurrentDocElemsToCheck(linkDoc);
            }
            else CheckLinkInst = null;
        }

        /// <summary>
        /// Ссылка на проверяемый документ
        /// </summary>
        public Document CheckDoc { get; }

        /// <summary>
        /// BoundingBoxXYZ для проверки
        /// </summary>
        public BoundingBoxXYZ CheckBBox { get; }

        /// <summary>
        /// Outline для проверки
        /// </summary>
        public Outline CheckOutline { get; }

        public Transform CheckDocTransform { get; }

        public XYZ CheckDocBasePntPosition { get; }

        public Transform LinkTransfrom { get; }

        public XYZ LinkBasePntPosition { get; }

        /// <summary>
        /// Ссылка на линк
        /// </summary>
        public RevitLinkInstance CheckLinkInst { get; private set; }

        /// <summary>
        /// Коллекция элементов, которые нужно детально проверить на коллизии (они попали в фильтры по геометрии CheckOutline)
        /// </summary>
        public HashSet<Element> CurrentDocElemsToCheck { get; } = new HashSet<Element>(new ElementComparerById());

        /// <summary>
        /// Получить коллекцию уточненных элементов по элементу на проверку ДЛЯ СВЯЗИ (для текущего документа этот список уже сформирован по элементу)
        /// </summary>
        /// <returns></returns>
        public Element[] GetPotentioalIntersectedElems_ForLink(Outline addedElemOutline)
        {
            List<Element> potentialIntersectedElems = new List<Element>();

            // CheckLinkInst null - когда в одном ревит несколько моделей одного проекта открыты
            if (CheckLinkInst != null && CheckLinkInst.IsValidObject)
            {
                Document checkDoc = CheckLinkInst.GetLinkDocument();
                Outline checkOutline = new Outline(LinkTransfrom.OfPoint(addedElemOutline.MinimumPoint), LinkTransfrom.OfPoint(addedElemOutline.MaximumPoint));

                BoundingBoxIntersectsFilter bboxIntersectFilter = new BoundingBoxIntersectsFilter(checkOutline, 0.1);
                BoundingBoxIsInsideFilter bboxInsideFilter = new BoundingBoxIsInsideFilter(checkOutline, 0.1);

                // Подготовка коллекции эл-в пересекаемых и внутри расширенного BoundingBoxXYZ
                potentialIntersectedElems.AddRange(CurrentDocElemsToCheck
                    .Where(e => bboxIntersectFilter.PassesFilter(checkDoc, e.Id)));

                potentialIntersectedElems.AddRange(CurrentDocElemsToCheck
                    .Where(e => bboxInsideFilter.PassesFilter(checkDoc, e.Id)));
            }

            return potentialIntersectedElems.Distinct().ToArray();
        }

        private void SetCurrentDocElemsToCheck(Document currentDoc)
        {
            Outline checkOutline = CheckLinkInst == null
                ? CheckOutline
                : new Outline(LinkTransfrom.OfPoint(CheckOutline.MinimumPoint), LinkTransfrom.OfPoint(CheckOutline.MaximumPoint));

            BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(checkOutline, 0.1);
            BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(checkOutline, 0.1);

            CurrentDocElemsToCheck.UnionWith(new FilteredElementCollector(currentDoc)
                .WherePasses(new LogicalAndFilter(ElemCatLogicalOrFilter, intersectsFilter))
                .Where(ElemExtraFilterFunc));

            CurrentDocElemsToCheck.UnionWith(new FilteredElementCollector(currentDoc)
                .WherePasses(new LogicalAndFilter(ElemCatLogicalOrFilter, insideFilter))
                .Where(ElemExtraFilterFunc));
        }
    }
}

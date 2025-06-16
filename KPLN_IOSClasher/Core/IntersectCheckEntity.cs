using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность области анализа на коллизии
    /// </summary>
    internal sealed class IntersectCheckEntity
    {
        private static LogicalOrFilter _elemCatLogicalOrFilter;
        private static Func<Element, bool> _elemFilterFunc;

        private double _checkDocBPElevation = -99.99;

        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        public static BuiltInCategory[] BuiltInCategories
        {
            get =>
                new BuiltInCategory[]
                { 
                    // ОВВК (ЭОМСС - огнезащита)
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    // ЭОМСС
                    BuiltInCategory.OST_CableTray,
                };
        }

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
                    catFilters.AddRange(BuiltInCategories.Select(bic => new ElementCategoryFilter(bic)));

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
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("Огнезащитный короб"));
                }

                return _elemFilterFunc;
            }
        }

        public IntersectCheckEntity(Document checkDoc, Outline filterOutline)
        {
            CheckDoc = checkDoc;
            CheckOutline = filterOutline;

            SetCurrentDocElemsToCheck(CheckDoc);
        }

        public IntersectCheckEntity(Document checkDoc, Outline filterOutline, RevitLinkInstance linkInst)
        {
            CheckDoc = checkDoc;
            CheckOutline = filterOutline;

            CheckLinkInst = linkInst;

            Document linkDoc = CheckLinkInst.GetLinkDocument();

            // Если открыто сразу несколько моделей одного проекта, то линки могут прилететь с другого файла. В таком случае - игнор и аннулирование CheckLinkInst
            if (linkDoc != null)
            {
                LinkTransfrom = GetLinkTransform(CheckLinkInst);
                
                SetCurrentDocElemsToCheck(linkDoc);
            }
            else CheckLinkInst = null;
        }

        /// <summary>
        /// Ссылка на проверяемый документ
        /// </summary>
        public Document CheckDoc { get; }

        /// <summary>
        /// Ссылка на отметку БТП проверяемого документа (только на неё идёт смещение)
        /// </summary>
        public double CheckDocBPElevation 
        {
            get
            {
                if(_checkDocBPElevation == -99.99)
                {
                    _checkDocBPElevation = new FilteredElementCollector(CheckDoc)
                        .OfClass(typeof(BasePoint))
                        .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                        .WhereElementIsNotElementType()
                        .FirstOrDefault()
                        .get_BoundingBox(null)
                        .Min
                        .Z;
                }
                _checkDocBPElevation = new FilteredElementCollector(CheckDoc)
                        .OfClass(typeof(BasePoint))
                        .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                        .WhereElementIsNotElementType()
                        .FirstOrDefault()
                        .get_BoundingBox(null)
                        .Min
                        .Z;

                return _checkDocBPElevation;
            }
        }

        /// <summary>
        /// Outline для проверки
        /// </summary>
        public Outline CheckOutline { get; }

        /// <summary>
        /// Трансформ для связи
        /// </summary>
        public Transform LinkTransfrom { get; }

        /// <summary>
        /// Ссылка на линк
        /// </summary>
        public RevitLinkInstance CheckLinkInst { get; private set; }

        /// <summary>
        /// Коллекция элементов, которые нужно детально проверить на коллизии (они попали в фильтры по геометрии CheckOutline)
        /// </summary>
        public HashSet<Element> CurrentDocElemsToCheck { get; } = new HashSet<Element>(new ElementComparerById());

        /// <summary>
        /// Задать Transform связи
        /// </summary>
        public static Transform GetLinkTransform(RevitLinkInstance linkInst)
        {
            Instance inst = linkInst as Instance;
            Transform instTrans = inst.GetTotalTransform();
            return instTrans;
        }

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
                Outline checkOutline = TransformFilterOutline_ToLink(addedElemOutline, LinkTransfrom);

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

        /// <summary>
        /// Трансформ для аутлайна, в координаты линка. Возможно опракидывание координат (например Y у MIN будет больше Y у MAX) - это
        /// не допустимо для создания фильтра
        /// </summary>
        private static Outline TransformFilterOutline_ToLink(Outline outlineToTransform, Transform linkTRansform)
        {
            // Inverse - т.к. возвращаюсь в координаты линка
            Outline transformOutline = new Outline(linkTRansform.Inverse.OfPoint(outlineToTransform.MinimumPoint), linkTRansform.Inverse.OfPoint(outlineToTransform.MaximumPoint));

            XYZ transOutlineMin = transformOutline.MinimumPoint;
            XYZ transOutlineMax = transformOutline.MaximumPoint;

            double minX = transOutlineMin.X;
            double minY = transOutlineMin.Y;
            double minZ = transOutlineMin.Z;

            double maxX = transOutlineMax.X;
            double maxY = transOutlineMax.Y;
            double maxZ = transOutlineMax.Z;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);
            double sminZ = Math.Min(minZ, maxZ);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);
            double smaxZ = Math.Max(minZ, maxZ);

            XYZ pntMin = new XYZ(sminX, sminY, sminZ);
            XYZ pntMax = new XYZ(smaxX, smaxY, smaxZ);

            return new Outline(pntMin, pntMax);
        }

        private void SetCurrentDocElemsToCheck(Document currentDoc)
        {
            Outline checkOutline = CheckLinkInst == null
                ? CheckOutline
                : TransformFilterOutline_ToLink(CheckOutline, LinkTransfrom);

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

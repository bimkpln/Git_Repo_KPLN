using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Shapes;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность области анализа на коллизии
    /// </summary>
    internal class IntersectCheckEntity
    {
        private static int[] _builtInCatIDs;

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

        /// <summary>
        /// Список id BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        public static int[] BuiltInCatIDs
        {
            get
            {
                if (_builtInCatIDs == null)
                    _builtInCatIDs = _builtInCategories.Select(bic => (int)bic).ToArray();

                return _builtInCatIDs;
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

            foreach (BuiltInCategory category in _builtInCategories)
            {
                CurrentDocElemsToCheck.UnionWith(new FilteredElementCollector(currentDoc)
                    .OfCategory(category)
                    .WherePasses(intersectsFilter)
                .ToElements());

                CurrentDocElemsToCheck.UnionWith(new FilteredElementCollector(currentDoc)
                    .OfCategory(category)
                    .WherePasses(insideFilter)
                    .ToElements());
            }
        }
    }
}

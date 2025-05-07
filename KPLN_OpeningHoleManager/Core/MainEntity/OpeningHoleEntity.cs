using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core.MainEntity
{
    public enum OpenigHoleShape
    {
        Rectangle,
        Circle
    }

    public class OpeningHoleEntity
    {
        private Solid _ohe_Solid;

        /// <summary>
        /// Ссылка на документ линка
        /// </summary>
        public Document OHE_LinkDocument { get; private protected set; }

        /// <summary>
        /// Ссылка на Transform для линка
        /// </summary>
        public Transform OHE_LinkTransform { get; private protected set; }

        /// <summary>
        /// Ссылка на элемент модели
        /// </summary>
        public Element OHE_Element { get; set; }

        /// <summary>
        /// Точка вставки элемента (ЗИ или отверстия)
        /// </summary>
        public XYZ OHE_Point { get; private protected set; }

        /// <summary>
        /// Имя параметра высоты отверстия/задания
        /// </summary>
        public string OHE_ParamNameHeight { get; private protected set; }

        /// <summary>
        /// Высота отверстия/задания
        /// </summary>
        public double OHE_Height { get; private protected set; }

        /// <summary>
        /// Имя параметра ширины отверстия/задания
        /// </summary>
        public string OHE_ParamNameWidth { get; private protected set; }

        /// <summary>
        /// Ширина отверстия/задания
        /// </summary>
        public double OHE_Width { get; private protected set; }

        /// <summary>
        /// Имя параметра радиуса отверстия/задания
        /// </summary>
        public string OHE_ParamNameRadius { get; private protected set; }

        /// <summary>
        /// Радиус отверстия/задания
        /// </summary>
        public double OHE_Radius { get; private protected set; }

        /// <summary>
        /// Имя параметра расширения отверстия/задания
        /// </summary>
        public string OHE_ParamNameExpander { get; private protected set; }

        /// <summary>
        /// Форма элемента (ЗИ или отверстия)
        /// </summary>
        public OpenigHoleShape OHE_Shape { get; private protected set; }

        /// <summary>
        /// Имя отдела
        /// </summary>
        public string OHE_SubDepartment_Name { get; private protected set; }

        /// <summary>
        /// Путь к семейству прямоугольного отверстия
        /// </summary>
        public string OHE_FamilyPath_Rectangle { get; private protected set; }

        /// <summary>
        /// Имя семейства прямоугольного отверстия
        /// </summary>
        public string OHE_FamilyName_Rectangle { get; private protected set; }

        /// <summary>
        /// Путь к семейству круглого отверстия
        /// </summary>
        public string OHE_FamilyPath_Circle { get; private protected set; }

        /// <summary>
        /// Имя семейства круглого отверстия
        /// </summary>
        public string OHE_FamilyName_Circle { get; private protected set; }

        /// <summary>
        /// Кэширование SOLID геометрии
        /// </summary>
        public Solid OHE_Solid
        {
            get
            {
                if (_ohe_Solid == null)
                {
                    if (OHE_LinkTransform == null)
                        _ohe_Solid = GeometryWorker.GetRevitElemSolid(OHE_Element);
                    else
                        _ohe_Solid = GeometryWorker.GetRevitElemSolid(OHE_Element, OHE_LinkTransform);
                }

                return _ohe_Solid;
            }

            protected private set => _ohe_Solid = value;
        }

        /// <summary>
        /// Задать форму отверстия по имени элементу
        /// </summary>
        public OpeningHoleEntity SetShapeByFamilyName(FamilyInstance fi)
        {
            string fiName = fi.Symbol.FamilyName;
            if (fiName.Equals(OHE_FamilyName_Rectangle))
                OHE_Shape = OpenigHoleShape.Rectangle;
            else if (fiName.Equals(OHE_FamilyName_Circle))
                OHE_Shape = OpenigHoleShape.Circle;
            else
                throw new Exception("Вы выбрали экзмеляр, который НЕ является заданием на отверстие");

            return this;
        }

        /// <summary>
        /// Задать Transform для линка
        /// </summary>
        /// <param name="inst">Instance связи</param>
        /// <returns></returns>
        public OpeningHoleEntity SetTransform(Instance inst)
        {
            Transform transform = inst.GetTransform();
            // Метка того, что базис трансформа тождество. Если нет, то создаём такой трансформ
            if (transform.IsTranslation)
                OHE_LinkTransform = transform;
            else
                OHE_LinkTransform = Transform.CreateTranslation(transform.Origin);

            return this;
        }

        /// <summary>
        /// Получить семейство отображающее коллизию
        /// </summary>
        private protected static FamilySymbol GetIntersectFamilySymbol(Document doc, string famPath, string famName, string famType)
        {
            Family mainFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(ft => ft.Name == famName);

            // Если в проекте нет - то грузим
            if (mainFamily == null)
            {
                bool result = doc.LoadFamily(famPath);
                if (!result)
                    throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                doc.Regenerate();

                // Повтор выборки, чтобы она не была пустой
                mainFamily = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(ft => ft.Name.Contains(famName));
            }

            FamilySymbol searchSymbol = null;
            // Поиск по разделу дока
            ISet<ElementId> fsIds = mainFamily.GetFamilySymbolIds();
            foreach (ElementId fsId in fsIds)
            {
                FamilySymbol fs = doc.GetElement(fsId) as FamilySymbol;
                if (fs.Name.Contains(famType))
                    searchSymbol = fs;
            }

            // Просто дефолтное значение
            if (searchSymbol == null)
                searchSymbol = doc.GetElement(mainFamily.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;

            searchSymbol.Activate();

            return searchSymbol;
        }
    }
}

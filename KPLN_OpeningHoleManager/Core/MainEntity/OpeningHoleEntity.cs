﻿using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_OpeningHoleManager.Core.MainEntity
{
    public enum OpenigHoleShape
    {
        Rectangular,
        Round
    }

    /// <summary>
    /// Обобщение отверстия в модели
    /// </summary>
    public class OpeningHoleEntity : InstanceGeomData
    {
        public OpeningHoleEntity(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Ссылка на документ линка
        /// </summary>
        public Document OHE_LinkDocument { get; private protected set; }

        /// <summary>
        /// Ссылка на Transform для линка
        /// </summary>
        public Transform OHE_LinkTransform { get; private protected set; }

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
        /// Задать форму отверстия по имени элементу
        /// </summary>
        public OpeningHoleEntity SetShapeByFamilyName(Element el)
        {
            if (el is FamilyInstance fi)
            {
                string fiName = fi.Symbol.FamilyName.ToLower();
                // Обрезаю из имени резеврные копии и копии семейств
                string clearedFiName = Regex.Replace(fiName, @"(\.\d+|\d+)$", "");
                if (clearedFiName.StartsWith(OHE_FamilyName_Rectangle.ToLower()))
                    OHE_Shape = OpenigHoleShape.Rectangular;
                else if (clearedFiName.StartsWith(OHE_FamilyName_Circle.ToLower()))
                    OHE_Shape = OpenigHoleShape.Round;
                else
                    throw new Exception("Вы выбрали экзмеляр, который НЕ является заданием на отверстие");
            }

            return this;
        }

        /// <summary>
        /// Задать Transform для линка
        /// </summary>
        /// <param name="inst">Instance связи</param>
        /// <returns></returns>
        public OpeningHoleEntity SetTransform(Instance inst)
        {
            OHE_LinkTransform = inst.GetTransform();

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

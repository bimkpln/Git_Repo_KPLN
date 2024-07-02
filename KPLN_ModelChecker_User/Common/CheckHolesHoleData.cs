using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по каждому отверстию АР
    /// </summary>
    internal class CheckHolesHoleData : CheckHolesEntity
    {
        private Face _mainHoleFace;

        public CheckHolesHoleData(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Коллекция элементов, которые пересекаются с CurrentHole
        /// </summary>
        public List<CheckHolesMEPData> IntesectElementsColl { get; private set; } = new List<CheckHolesMEPData>();

        /// <summary>
        /// Суммарная площадь пересекаемых элементов ИОС с CurrentHole
        /// </summary>
        public double SumIntersectArea { get; private set; }

        /// <summary>
        /// Основная поверхность отверстия (та, которая пересекается с элементами ИОС)
        /// </summary>
        public Face MainHoleFace
        {
            get
            {
                if (_mainHoleFace == null)
                    _mainHoleFace = GetMainFace(this.CurrentSolid);
                return _mainHoleFace;
            }
        }

        /// <summary>
        ///  Поиск основной поверхности отверстия
        /// </summary>
        private Face GetMainFace(Solid solid)
        {
            FaceArray faces = solid.Faces;
            foreach (Face face in faces)
            {
                PlanarFace planarFace = face as PlanarFace;
                if (planarFace != null && planarFace.XVector.Z == 1) return face;
            }
            
            throw new Exception($"Ошибка с получением основной поверхности отверстия или результата пересечений у отверстия с id: {this.CurrentElement.Id}. Отправь разработчику!");
        }

        /// <summary>
        /// Задать коллецию элементов и рассчитать площадь пересечения элементов, которые пересекаются с отверстием
        /// </summary>
        /// <param name="mepElements">Коллекция элементов ИОС, которые нужно проверить на пересечение</param>
        public void SetIntersectsData(List<CheckHolesMEPData> mepElements, List<CheckCommandError> notCriticalErrorElemColl)
        {
            XYZ currentCentroid = this.CurrentSolid.ComputeCentroid();
            foreach (CheckHolesMEPData mepData in mepElements)
            {
                foreach (XYZ locPoint in mepData.CurrentLocationColl)
                {
                    if (locPoint.DistanceTo(currentCentroid) < 30)
                    {
                        mepData.SetGeometryData(ViewDetailLevel.Fine);
                        if (mepData.CurrentSolid == null) 
                            continue;
                        try
                        {
                            Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(this.CurrentSolid, mepData.CurrentSolid, BooleanOperationsType.Intersect);
                            if (intersectionSolid != null && intersectionSolid.Volume > 0)
                            {
                                this.IntesectElementsColl.Add(mepData);
                                Face mainIntersectFace = GetMainFace(intersectionSolid);
                                SumIntersectArea += mainIntersectFace.Area;
                            }
                            break;
                        }
                        catch (Exception)
                        {
                            // Это ошибка создания солида. Возникает крайне редко, и такой элемент игнарирую,
                            // т.к. он все равно попадет в отчет (либо площади не хватит, либо будет незанятое отверстие)
                            break;
                        }
                    }
                }
            }
        }
    }
}

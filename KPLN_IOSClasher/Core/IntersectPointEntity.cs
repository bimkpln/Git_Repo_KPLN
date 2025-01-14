using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность для запаковки данных по точкам пересечения для файла
    /// </summary>
    internal sealed class IntersectPointEntity
    {
        public IntersectPointEntity(
            XYZ pnt,
            int addedElement_Id,
            int oldElement_Id,
            int linkId,
            DBUser user)
        {
            IntersectPoint = pnt;

            AddedElement_Id = addedElement_Id;
            OldElement_Id = oldElement_Id;

            LinkInstance_Id = linkId;

            CurrentUser = user;
        }

        /// <summary>
        /// Точка пересечения
        /// </summary>
        public XYZ IntersectPoint { get; }

        /// <summary>
        /// Id нового/измененного элемента
        /// </summary>
        public int AddedElement_Id { get; }

        /// <summary>
        /// Id уже присутсвующего элемента
        /// </summary>
        public int OldElement_Id { get; }


        /// <summary>
        /// Ссылка на id элемента связи (если линка нет, то "-1")
        /// </summary>
        public int LinkInstance_Id { get; }

        /// <summary>
        /// Данные по текущему пользователю
        /// </summary>
        public DBUser CurrentUser { get; }
    }

    /// <summary>
    /// Класс для сравнения PointEntityIntesectCheck по XYZ (для создания HashSet)
    /// </summary>
    internal sealed class PointEntityComparerByXYZ : IEqualityComparer<IntersectPointEntity>
    {
        private const double _tolerance = 0.5;

        public bool Equals(IntersectPointEntity x, IntersectPointEntity y)
        {
            if (x.IntersectPoint == null || y.IntersectPoint == null)
                return false;

            return Math.Abs(x.IntersectPoint.DistanceTo(y.IntersectPoint)) > _tolerance;
        }

        public int GetHashCode(IntersectPointEntity obj)
        {
            if (obj.IntersectPoint == null)
                return 0;

            // Заакругленне каардынат да хібнасці для стварэння стабільнага хэш-кода
            int hashX = (int)(obj.IntersectPoint.X / _tolerance);
            int hashY = (int)(obj.IntersectPoint.Y / _tolerance);
            int hashZ = (int)(obj.IntersectPoint.Z / _tolerance);

            // Камбінаванне хэш-кодаў каардынат
            return hashX ^ (hashY << 2) ^ (hashZ >> 2);
        }
    }
}

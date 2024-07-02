using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common
{
    internal class PntComparer : IComparer<XYZ>
    {
        private readonly XYZ _centerPoint;

        public PntComparer(XYZ centerPoint)
        {
            _centerPoint = centerPoint;
        }

        public int Compare(XYZ pointA, XYZ pointB)
        {
            if (pointA.X - _centerPoint.X >= 0 && pointB.X - _centerPoint.X < 0)
                return 1;
            if (pointA.X - _centerPoint.X < 0 && pointB.X - _centerPoint.X >= 0)
                return -1;

            if (pointA.X - _centerPoint.X == 0 && pointB.X - _centerPoint.X == 0)
            {
                if (pointA.Y - _centerPoint.Y >= 0 || pointB.Y - _centerPoint.Y >= 0)
                    if (pointA.Y > pointB.Y)
                        return 1;
                    else return -1;
                if (pointB.Y > pointA.Y)
                    return 1;
                else return -1;
            }

            // compute the cross product of vectors (CenterPoint -> a) x (CenterPoint -> b)
            double det = (pointA.X - _centerPoint.X) * (pointB.Y - _centerPoint.Y) -
                             (pointB.X - _centerPoint.X) * (pointA.Y - _centerPoint.Y);
            if (det < 0)
                return 1;
            if (det > 0)
                return -1;

            // points a and b are on the same line from the CenterPoint
            // check which point is closer to the CenterPoint
            double d1 = (pointA.X - _centerPoint.X) * (pointA.X - _centerPoint.X) +
                            (pointA.Y - _centerPoint.Y) * (pointA.Y - _centerPoint.Y);
            double d2 = (pointB.X - _centerPoint.X) * (pointB.X - _centerPoint.X) +
                            (pointB.Y - _centerPoint.Y) * (pointB.Y - _centerPoint.Y);
            if (d1 > d2)
                return 1;
            else return -1;
        }
    }
}

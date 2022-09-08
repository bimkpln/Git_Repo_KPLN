using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    internal class ClockwiseComparerTool : IComparer<XYZ>
    {
        private XYZ CenterPoint;

        public ClockwiseComparerTool(XYZ centerPoint)
        {
            CenterPoint = centerPoint;
        }

        public int Compare(XYZ pointA, XYZ pointB)
        {
            if (pointA.X - CenterPoint.X >= 0 && pointB.X - CenterPoint.X < 0)
                return 1;
            if (pointA.X - CenterPoint.X < 0 && pointB.X - CenterPoint.X >= 0)
                return -1;

            if (pointA.X - CenterPoint.X == 0 && pointB.X - CenterPoint.X == 0)
            {
                if (pointA.Y - CenterPoint.Y >= 0 || pointB.Y - CenterPoint.Y >= 0)
                    if (pointA.Y > pointB.Y)
                        return 1;
                    else return -1;
                if (pointB.Y > pointA.Y)
                    return 1;
                else return -1;
            }

            // compute the cross product of vectors (CenterPoint -> a) x (CenterPoint -> b)
            double det = (pointA.X - CenterPoint.X) * (pointB.Y - CenterPoint.Y) -
                             (pointB.X - CenterPoint.X) * (pointA.Y - CenterPoint.Y);
            if (det < 0)
                return 1;
            if (det > 0)
                return -1;

            // points a and b are on the same line from the CenterPoint
            // check which point is closer to the CenterPoint
            double d1 = (pointA.X - CenterPoint.X) * (pointA.X - CenterPoint.X) +
                            (pointA.Y - CenterPoint.Y) * (pointA.Y - CenterPoint.Y);
            double d2 = (pointB.X - CenterPoint.X) * (pointB.X - CenterPoint.X) +
                            (pointB.Y - CenterPoint.Y) * (pointB.Y - CenterPoint.Y);
            if (d1 > d2)
                return 1;
            else return -1;
        }
    }
}

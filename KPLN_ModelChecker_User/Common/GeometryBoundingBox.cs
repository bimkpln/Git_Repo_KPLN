using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Common
{
    public class GeometryBoundingBox
    {
        public BoundingBoxXYZ BoundingBox 
        {
            get
            {
                BoundingBoxXYZ box = new BoundingBoxXYZ();
                box.Min = Min;
                box.Max = Max;
                return box;
            }
        }
        public XYZ Min { get; private set; }
        public XYZ Max { get; private set; }
        public XYZ Centroid
        {
            get
            {
                return new XYZ((Max.X + Min.X) / 2, (Max.Y + Min.Y) / 2, (Max.Z + Min.Z) / 2);
            }
        }
        public GeometryBoundingBox(Solid solid)
        {
            BoundingBoxXYZ box = solid.GetBoundingBox();
            Min = box.Min + solid.ComputeCentroid();
            Max = box.Max + solid.ComputeCentroid();
        }
        public void AppendSolid(Solid solid)
        {
            BoundingBoxXYZ box = solid.GetBoundingBox();
            XYZ centroid = solid.ComputeCentroid();
            Min = new XYZ(Math.Min(box.Min.X + centroid.X, Min.X), Math.Min(box.Min.Y + centroid.Y, Min.Y), Math.Min(box.Min.Z + centroid.Z, Min.Z));
            Max = new XYZ(Math.Max(box.Max.X + centroid.X, Max.X), Math.Max(box.Max.Y + centroid.Y, Max.Y), Math.Max(box.Max.Z + centroid.Z, Max.Z));
        }
    }
}

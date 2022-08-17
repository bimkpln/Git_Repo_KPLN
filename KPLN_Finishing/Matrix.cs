using Autodesk.Revit.DB;
using KPLN_Finishing.Forms;
using System;
using System.Collections.Generic;

namespace KPLN_Finishing
{
    public class MatrixElement
    {
        public Element Element { get; }
        public MetaRoom MetaRoom { get; }
        public Solid Solid { get; }
        public XYZ Centroid { get; }
        public BoundingBoxXYZ BoundingBox { get; }
        public LocationCurve Curve { get; }
        public MatrixElement(Element element, Solid solid, XYZ centroid, BoundingBoxXYZ boundingBox, LocationCurve curve = null)
        {
            this.Element = element;
            this.Solid = solid;
            this.Centroid = centroid;
            this.BoundingBox = boundingBox;
            this.Curve = curve;
            if (element == null || solid == null || centroid == null || boundingBox == null)
            {
                throw new Exception("Параметры конструктора не могут быть пустыми");
            }
        }
        public MatrixElement(MetaRoom room, Solid solid, XYZ centroid, BoundingBoxXYZ boundingBox, LocationCurve curve = null)
        {
            this.MetaRoom = room;
            this.Element = room.Room;
            this.Solid = solid;
            this.Centroid = centroid;
            this.BoundingBox = boundingBox;
            this.Curve = curve;
            if (room == null || solid == null || centroid == null || boundingBox == null)
            {
                throw new Exception("Параметры конструктора не могут быть пустыми");
            }
        }
    }
    public class MatrixContainer
    {
        private static int AmountMax = 0;
        private static int AmountCurrent = 0;
        public double Length { get; }
        public double Width { get; }
        public double Height { get; }
        public List<MatrixContainer> SubContainers = new List<MatrixContainer>();
        public List<MatrixElement> SubElements = new List<MatrixElement>();
        public BoundingBoxXYZ BoundingBox { get; set; }
        public MatrixContainer(XYZ min, XYZ max, MatrixContainer parent = null)
        {
            AmountMax += 1;
            this.BoundingBox = new BoundingBoxXYZ();
            this.BoundingBox.Min = min;
            this.BoundingBox.Max = max;
            this.Width = Math.Round(Math.Abs(this.BoundingBox.Max.X - this.BoundingBox.Min.X));
            this.Length = Math.Round(Math.Abs(this.BoundingBox.Max.Y - this.BoundingBox.Min.Y));
            this.Height = Math.Round(Math.Abs(this.BoundingBox.Max.Z - this.BoundingBox.Min.Z));
            this.SubContainers = new List<MatrixContainer>();
            if (this.Length >= 22.0 && this.Height >= 22.0 && this.Width >= 22.0)
            {
                AmountMax += 8;
                //Creating substructure
                //Note: Parent BoundingBoxXYZ Min & Max need to match condition: x,y,z % 2 = 0
                //BOT BOXES
                //##
                //$#
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X, this.BoundingBox.Min.Y, this.BoundingBox.Min.Z),
                    new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height / 2), this));
                //$#
                //##
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z),
                    new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length, this.BoundingBox.Min.Z + this.Height / 2), this));
                //##
                //#$
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y, this.BoundingBox.Min.Z),
                    new XYZ(this.BoundingBox.Min.X + this.Width, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height / 2), this));
                //#$
                //##
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z),
                    new XYZ(this.BoundingBox.Min.X + this.Width, this.BoundingBox.Min.Y + this.Length, this.BoundingBox.Min.Z + this.Height / 2), this));
                //TOP BOXES
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X, this.BoundingBox.Min.Y, this.BoundingBox.Min.Z + this.Height / 2),
                    new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height), this));
                //$#
                //##
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height / 2),
                    new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length, this.BoundingBox.Min.Z + this.Height), this));
                //##
                //#$
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y, this.BoundingBox.Min.Z + this.Height / 2),
                    new XYZ(this.BoundingBox.Min.X + this.Width, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height), this));
                //#$
                //##
                this.SubContainers.Add(new MatrixContainer(new XYZ(this.BoundingBox.Min.X + this.Width / 2, this.BoundingBox.Min.Y + this.Length / 2, this.BoundingBox.Min.Z + this.Height / 2),
                    new XYZ(this.BoundingBox.Min.X + this.Width, this.BoundingBox.Min.Y + this.Length, this.BoundingBox.Min.Z + this.Height), this));
            }
        }
        protected bool IsEmpty()
        {
            if (this.SubContainers.Count == 0)
            {
                if (this.SubElements.Count == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                foreach (MatrixContainer subContainer in this.SubContainers)
                {
                    if (!subContainer.IsEmpty())
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        public void Optimize(bool reset = false)
        {
            if (reset)
            {
                AmountCurrent = 0;
            }
            List<MatrixContainer> newSubContainers = new List<MatrixContainer>();
            foreach (MatrixContainer subContainer in this.SubContainers)
            {
                if (!subContainer.IsEmpty())
                {
                    newSubContainers.Add(subContainer);
                    AmountCurrent += 1;
                }
            }
            this.SubContainers = newSubContainers;
            if (this.SubContainers.Count != 0)
            {
                foreach (MatrixContainer subContainer in this.SubContainers)
                {
                    subContainer.Optimize();
                }
            }
        }
        public void InsertItem(MatrixElement element)
        {
            if (this.SubContainers.Count != 0)
            {
                if (Tools.IntersectsBoundingBox(element.BoundingBox, this.BoundingBox))
                {
                    foreach (MatrixContainer container in this.SubContainers)
                    {
                        container.InsertItem(element);
                    }
                }
            }
            else
            {
                if (Tools.IntersectsBoundingBox(element.BoundingBox, this.BoundingBox))
                {
                    this.SubElements.Add(element);
                }
            }
        }
        public List<MatrixElement> GetItems(MatrixElement element)
        {
            List<int> elementIdList = new List<int>();
            List<MatrixElement> elementList = new List<MatrixElement>();
            if (Tools.IntersectsBoundingBox(element.BoundingBox, this.BoundingBox))
            {
                foreach (MatrixContainer subcontainer in this.SubContainers)
                {
                    List<MatrixElement> subelementList = subcontainer.GetItems(element);
                    if (subelementList.Count != 0)
                    {
                        foreach (MatrixElement subelement in subelementList)
                        {
                            int elementId = subelement.Element.Id.IntegerValue;
                            if (!elementIdList.Contains(elementId))
                            {
                                elementList.Add(subelement);
                                elementIdList.Add(subelement.Element.Id.IntegerValue);
                            }
                        }
                    }
                }
            }
            return elementList;
        }
        public string GetSubMatrix(int i = 0)
        {
            int next = i + 1;
            string res = "\n";
            int ticker = 0;
            while (ticker != i)
            {
                ticker++;
                res += "    ";

            }
            res += string.Format("{0} : {1}", this.BoundingBox.Min.ToString(), this.BoundingBox.Max.ToString());
            if (this.SubElements.Count != 0)
            {
                res += string.Format(" - {0}", this.SubElements.Count.ToString());
            }
            foreach (MatrixContainer subMc in this.SubContainers)
            {
                res += subMc.GetSubMatrix(next);
            }
            return res;
        }
        public List<MatrixElement> GetContext(MatrixElement element)
        {
            List<MatrixElement> listElements = new List<MatrixElement>();
            List<int> listId = new List<int>();
            foreach (MatrixContainer container in this.SubContainers)
            {
                if (Tools.IntersectsBoundingBox(element.BoundingBox, container.BoundingBox))
                {
                    foreach (MatrixElement matrixElement in container.GetElements(element))
                    {
                        if (!listId.Contains(matrixElement.Element.Id.IntegerValue))
                        {
                            listElements.Add(matrixElement);
                            listId.Add(matrixElement.Element.Id.IntegerValue);
                        }
                    }
                }
            }
            return listElements;
        }
        public List<MatrixElement> GetGlobalContext(MatrixElement element)
        {
            List<MatrixElement> listElements = new List<MatrixElement>();
            List<int> listId = new List<int>();
            foreach (MatrixContainer container in this.SubContainers)
            {
                if (Tools.IntersectsOrNearBoundingBox(element.BoundingBox, container.BoundingBox))
                {
                    foreach (MatrixElement matrixElement in container.GetElements(element))
                    {
                        if (!listId.Contains(matrixElement.Element.Id.IntegerValue))
                        {
                            listElements.Add(matrixElement);
                            listId.Add(matrixElement.Element.Id.IntegerValue);
                        }
                    }
                }
            }
            return listElements;
        }
        private List<MatrixElement> GetElements(MatrixElement element)
        {
            List<MatrixElement> listElements = new List<MatrixElement>();
            if (this.SubElements.Count != 0)
            {
                foreach (MatrixElement el in this.SubElements)
                {
                    listElements.Add(el);
                }
            }
            if (this.SubContainers.Count != 0)
            {
                foreach (MatrixContainer container in this.SubContainers)
                {
                    if (Tools.IntersectsBoundingBox(element.BoundingBox, container.BoundingBox))
                    {
                        foreach (MatrixElement el in container.GetElements(element))
                        {
                            listElements.Add(el);
                        }
                    }
                }
            }
            return listElements;
        }
    }
    public class Matrix
    {
        public BoundingBoxXYZ BoundingBox = new BoundingBoxXYZ();
        private List<MatrixElement> Elements = new List<MatrixElement>();
        private List<MatrixElement> Rooms = new List<MatrixElement>();
        private MatrixContainer Container { get; set; }
        public void Clear()
        {
            Elements.Clear();
            Rooms.Clear();
            BoundingBox.Max = new XYZ(-999, -999, -999);
            BoundingBox.Min = new XYZ(999, 999, 999);
        }
        public void AppendToElements(MatrixElement element)
        {
            Elements.Add(element);
            GlobalXYZ minPoint = new GlobalXYZ(new XYZ(BoundingBox.Min.X, BoundingBox.Min.Y, BoundingBox.Min.Z));
            GlobalXYZ maxPoint = new GlobalXYZ(new XYZ(BoundingBox.Max.X, BoundingBox.Max.Y, BoundingBox.Max.Z));
            if (maxPoint.X < element.BoundingBox.Max.X)
            {
                maxPoint.X = element.BoundingBox.Max.X;
            }
            if (maxPoint.Y < element.BoundingBox.Max.Y)
            {
                maxPoint.Y = element.BoundingBox.Max.Y;
            }
            if (maxPoint.Z < element.BoundingBox.Max.Z)
            {
                maxPoint.Z = element.BoundingBox.Max.Z;
            }
            if (minPoint.X > element.BoundingBox.Min.X)
            {
                minPoint.X = element.BoundingBox.Min.X;
            }
            if (minPoint.Y > element.BoundingBox.Min.Y)
            {
                minPoint.Y = element.BoundingBox.Min.Y;
            }
            if (minPoint.Z > element.BoundingBox.Min.Z)
            {
                minPoint.Z = element.BoundingBox.Min.Z;
            }
            BoundingBox.Min = minPoint.Floor().GetPoint();
            BoundingBox.Max = maxPoint.Ceil().GetPoint();
        }
        public void AppendToRooms(MatrixElement element)
        {
            Rooms.Add(element);
            GlobalXYZ minPoint = new GlobalXYZ(BoundingBox.Min);
            GlobalXYZ maxPoint = new GlobalXYZ(BoundingBox.Max);
            if (maxPoint.X < element.BoundingBox.Max.X)
            {
                maxPoint.X = element.BoundingBox.Max.X;
            }
            if (maxPoint.Y < element.BoundingBox.Max.Y)
            {
                maxPoint.Y = element.BoundingBox.Max.Y;
            }
            if (maxPoint.Z < element.BoundingBox.Max.Z)
            {
                maxPoint.Z = element.BoundingBox.Max.Z;
            }
            if (minPoint.X > element.BoundingBox.Min.X)
            {
                minPoint.X = element.BoundingBox.Min.X;
            }
            if (minPoint.Y > element.BoundingBox.Min.Y)
            {
                minPoint.Y = element.BoundingBox.Min.Y;
            }
            if (minPoint.Z > element.BoundingBox.Min.Z)
            {
                minPoint.Z = element.BoundingBox.Min.Z;
            }
            BoundingBox.Min = minPoint.GetPoint();
            BoundingBox.Max = maxPoint.GetPoint();
        }
        public MatrixContainer PrepareContainers()
        {
            double maxLength = Math.Max(Math.Round(Math.Abs(BoundingBox.Max.X - BoundingBox.Min.X)), Math.Round(Math.Abs(BoundingBox.Max.Y - BoundingBox.Min.Y)));
            maxLength = Math.Max(maxLength, Math.Round(Math.Abs(BoundingBox.Max.Z - BoundingBox.Min.Z)));
            Container = new MatrixContainer(BoundingBox.Min, BoundingBox.Max);
            return Container;
        }
        public Matrix()
        {
            Clear();
        }
    }
    public class GlobalXYZ
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public GlobalXYZ Ceil()
        {
            this.X = Math.Ceiling(Math.Ceiling(this.X) / 22) * 22;
            this.Y = Math.Ceiling(Math.Ceiling(this.Y) / 22) * 22;
            this.Z = Math.Ceiling(Math.Ceiling(this.Z) / 22) * 22;
            return this;
        }
        public GlobalXYZ Floor()
        {
            this.X = Math.Floor(Math.Floor(this.X) / 22) * 22;
            this.Y = Math.Floor(Math.Floor(this.Y) / 22) * 22;
            this.Z = Math.Floor(Math.Floor(this.Z) / 22) * 22;
            return this;
        }
        public GlobalXYZ(double x, double y, double z)
        {
            this.X = x;
            this.Y = y;
            this.Z = z;
        }
        public GlobalXYZ(XYZ point)
        {
            this.X = point.X;
            this.Y = point.Y;
            this.Z = point.Z;
        }
        public XYZ GetPoint()
        {
            return new XYZ(this.X, this.Y, this.Z);
        }
    }
}

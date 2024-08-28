using Autodesk.Revit.DB;

namespace KPLN_Tools.Common.SS_System
{
    /// <summary>
    /// Обертка для закрытого класса Revit API - GraphicsStyle. Нужен для подгрузки в окно (переопределение Equals и GetHashCode)
    /// </summary>
    public class GraphicsStyleWrapper
    {
        public GraphicsStyle RevitGraphicsStyle { get; }

        public GraphicsStyleWrapper(GraphicsStyle style)
        {
            RevitGraphicsStyle = style;
        }

        public override bool Equals(object obj)
        {
            if (obj is GraphicsStyleWrapper otherWrapper)
            {
                return RevitGraphicsStyle.Id == otherWrapper.RevitGraphicsStyle.Id;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return RevitGraphicsStyle.Id.GetHashCode();
        }

        public override string ToString()
        {
            return RevitGraphicsStyle.Name;
        }
    }
}

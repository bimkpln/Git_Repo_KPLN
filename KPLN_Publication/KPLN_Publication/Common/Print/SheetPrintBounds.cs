namespace KPLN_Publication
{
    /// <summary>Габариты в координатах листа (внутренние единицы Revit).</summary>
    internal sealed class SheetPrintBounds
    {
        internal const double Tolerance = 1.0 / 304.8;

        internal double MinX { get; }
        internal double MinY { get; }
        internal double MaxX { get; }
        internal double MaxY { get; }

        internal SheetPrintBounds(double minX, double minY, double maxX, double maxY)
        {
            MinX = minX;
            MinY = minY;
            MaxX = maxX;
            MaxY = maxY;
        }

        internal bool ExtendsOutside(SheetPrintBounds frame)
        {
            return MinX < frame.MinX - Tolerance || MinY < frame.MinY - Tolerance
                || MaxX > frame.MaxX + Tolerance || MaxY > frame.MaxY + Tolerance;
        }
    }
}

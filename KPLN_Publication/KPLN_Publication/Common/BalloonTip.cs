namespace KPLN_Publication
{
    public static class BalloonTip
    {
        public static void Show(string title, string message)
        {
            Autodesk.Internal.InfoCenter.ResultItem ri = new Autodesk.Internal.InfoCenter.ResultItem();

            ri.Category = title;
            ri.Title = message;
            //ri.TooltipText = tolltip;

            ri.Type = Autodesk.Internal.InfoCenter.ResultType.LocalFile;

            ri.IsFavorite = true;
            ri.IsNew = true;

            Autodesk.Windows.ComponentManager
                .InfoCenterPaletteManager.ShowBalloon(ri);
        }
    }
}

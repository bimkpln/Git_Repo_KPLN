using Autodesk.Revit.DB;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для рабочего набора
    /// </summary>
    public sealed class WSEntity
    {
        public WSEntity(Workset ws)
        {
            RevitWS = ws;

            RevitWSId = ws.Id;
            RevitWSName = ws.Name;
        }

        public Workset RevitWS { get; }

        public WorksetId RevitWSId { get; }

        public string RevitWSName { get; }
    }
}

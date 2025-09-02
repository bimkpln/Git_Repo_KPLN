using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_Lib.Forms.Entities
{
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
     
        public bool IsSelected { get; set; }
    }
}

using Autodesk.Revit.DB;
using System.Collections.ObjectModel;

namespace KPLN_Tools.Common
{
    public class MonitorParamRule
    {
        public ObservableCollection<Parameter> DocParamColl { get; set; }

        public ObservableCollection<Parameter> LinkParamColl { get; set; }

        public MonitorParamRule(ObservableCollection<Parameter> docParamColl, ObservableCollection<Parameter> linkParamColl)
        {
            DocParamColl = docParamColl;
            LinkParamColl = linkParamColl;
        }

        public Parameter SelectedSourceParameter { get; set; }

        public Parameter SelectedTargetParameter { get; set; }
    }
}
using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.CheckParam.Builder
{
    internal class AuditBuilder_AR : AbstrAuditBuilder
    {
        public AuditBuilder_AR(Document doc, string docMainTitle) : base(doc, docMainTitle)
        {
        }

        public override void Prepare()
        {
            throw new NotImplementedException();
        }

        public override bool ExecuteParamsAudit(ProgressInfoViewModel pb)
        {
            throw new NotImplementedException();
        }
    }
}

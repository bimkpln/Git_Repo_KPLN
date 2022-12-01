using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator
{
    public class StubOutput : Output
    {
        public override void PrintDebug(string value, OutputMessageType type, bool check)
        {
            return;
        }

        public override void PrintErr(Exception e, string value)
        {
            return;
        }

        public override void PrintInfo(string value, OutputMessageType type)
        {
            TaskDialog.Show("Отчёт", value);
        }
    }
}

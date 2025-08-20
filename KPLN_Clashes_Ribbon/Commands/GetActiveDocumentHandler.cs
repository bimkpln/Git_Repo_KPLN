using Autodesk.Revit.UI;

namespace KPLN_Clashes_Ribbon.Commands
{
    public sealed class GetActiveDocumentHandler : IExternalEventHandler
    {
        public UIApplication ResultUIApplication { get; private set; }

        public void Execute(UIApplication app)
        {
            ResultUIApplication = app;
        }

        public string GetName()
        {
            return "GetActiveDocumentHandler";
        }
    }
}

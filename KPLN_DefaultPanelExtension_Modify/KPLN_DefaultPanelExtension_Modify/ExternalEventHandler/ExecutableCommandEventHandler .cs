using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;

namespace KPLN_DefaultPanelExtension_Modify.ExternalEventHandler
{
    internal class ExecutableCommandEventHandler : IExternalEventHandler
    {
        private readonly Queue<IExecutableCommand> _commandQueue;

        public ExecutableCommandEventHandler(Queue<IExecutableCommand> commandQueue)
        {
            _commandQueue = commandQueue;
        }

        public void Execute(UIApplication app)
        {
            while (_commandQueue.Count != 0)
            {
                _commandQueue.Dequeue().Execute(app);
            }
        }

        public string GetName() => nameof(ExecutableCommandEventHandler);
    }
}

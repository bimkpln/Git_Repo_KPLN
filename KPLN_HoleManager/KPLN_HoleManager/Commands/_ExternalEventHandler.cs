using System;
using Autodesk.Revit.UI;

namespace KPLN_HoleManager.Commands
{
    public class _ExternalEventHandler : IExternalEventHandler
    {
        private static ExternalEvent _externalEvent;
        private static _ExternalEventHandler _instance;
        private Action<UIApplication> _action;

        private _ExternalEventHandler() { }

        public static _ExternalEventHandler Instance
        {
            get
            {
                if (_instance == null)
                {
                    throw new InvalidOperationException("ExternalEventHandler не инициализирован." +
                        "\nВызовите Initialize() в контексте API Revit.");
                }
                return _instance;
            }
        }

        public static void Initialize()
        {
            if (_instance == null)
            {
                _instance = new _ExternalEventHandler();
                _externalEvent = ExternalEvent.Create(_instance);
            }
        }

        public void Raise(Action<UIApplication> action)
        {
            _action = action;
            _externalEvent.Raise();
        }

        public void Execute(UIApplication app)
        {
            _action?.Invoke(app);
        }

        public string GetName() => "External Event Handler";
    }
}
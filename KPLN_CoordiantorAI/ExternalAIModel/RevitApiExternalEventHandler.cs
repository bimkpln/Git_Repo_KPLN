using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.ExternalModel;

namespace KPLN_CoordiantorAI.ExternalAIModel
{

    //Этот класс нужен чтобы реализовать асинхронный вызов Revit API через ExternalEvent — механизм, 
    //который позволяет выполнять операции с Revit API из фонового потока(например, из WPF-окна), 
    //не блокируя интерфейс Revit. Ставит в очередь задачу на выполнение в главном потоке Revit и получает результат асинхронно

    internal class RevitApiExternalEventHandler : IExternalEventHandler
    {

        private readonly object _syncRoot = new object();
        private readonly ExternalEvent _externalEvent;
        private PendingRequest _pendingRequest;

        public RevitApiExternalEventHandler()
        {
            _externalEvent = ExternalEvent.Create(this);
        }

        public Task<object> SetViewSectionBoxToElementsAsync(
            Document document,
            UIDocument uiDocument,
            List<int> elementIds,
            double marginMM = 500)
        {
            PendingRequest request = new PendingRequest
            {
                CommandName = "set_view_section_box_to_elements",
                Document = document,
                UiDocument = uiDocument,
                ElementIds = elementIds == null ? new List<int>() : new List<int>(elementIds),
                MarginMM = marginMM,
                Completion = new TaskCompletionSource<object>()
            };

            lock (_syncRoot)
            {
                if (_pendingRequest != null)
                {
                    request.Completion.SetResult(new
                    {
                        success = false,
                        error = "Revit уже выполняет другую команду из окна ИИ. Повторите запрос через несколько секунд."
                    });

                    return request.Completion.Task;
                }

                _pendingRequest = request;
            }

            try
            {
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                ClearPendingRequest(request);
                request.Completion.SetResult(new
                {
                    success = false,
                    error = "Не удалось передать команду в Revit API context: " + ex.Message
                });
            }

            return request.Completion.Task;
        }

        public void Execute(UIApplication app)
        {
            PendingRequest request = TakePendingRequest();
            if (request == null)
                return;

            try
            {
                object result;
                switch (request.CommandName)
                {
                    case "set_view_section_box_to_elements":
                        result = Commands.SetViewSectionBoxToElements(
                            request.Document,
                            request.UiDocument,
                            request.ElementIds,
                            request.MarginMM);
                        break;

                    default:
                        result = new
                        {
                            success = false,
                            error = "Неизвестная команда Revit ExternalEvent: " + request.CommandName
                        };
                        break;
                }

                request.Completion.SetResult(result);
            }
            catch (Exception ex)
            {
                request.Completion.SetResult(new
                {
                    success = false,
                    error = "Ошибка при выполнении команды в Revit API context: " + ex.Message
                });
            }
        }

        public string GetName()
        {
            return "KPLN Coordinator AI Revit API Command Handler";
        }

        private PendingRequest TakePendingRequest()
        {
            lock (_syncRoot)
            {
                PendingRequest request = _pendingRequest;
                _pendingRequest = null;
                return request;
            }
        }

        private void ClearPendingRequest(PendingRequest request)
        {
            lock (_syncRoot)
            {
                if (ReferenceEquals(_pendingRequest, request))
                    _pendingRequest = null;
            }
        }

        private sealed class PendingRequest
        {
            public string CommandName { get; set; }
            public Document Document { get; set; }
            public UIDocument UiDocument { get; set; }
            public List<int> ElementIds { get; set; }
            public double MarginMM { get; set; }
            public TaskCompletionSource<object> Completion { get; set; }
        }

    }
}

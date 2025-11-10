using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class UnsubscribeHandler : IExternalEventHandler
    {
#if Debug2020 || Revit2020
        НЕТ ТАКОГО КЛАССА. 
        НУЖЕН КОСТЫЛЬ ПО ВЫБОРКЕ (ОНО НАМ НУЖНО??? CHATGPT ПРЕДЛАГАЕТ ЧЕРЕЗ Idling - ЭТО СЛИШКОМ, НО ПОХОЖЕ ЭТО ЕДИНСТВЕННОЕ РЕШЕНИЕ)
        public EventHandler<SelectionChangedEventArgs> Handler { get; set; }
#else
        public EventHandler<SelectionChangedEventArgs> Handler { get; set; }
#endif


        public void Execute(UIApplication app)
        {
            app.SelectionChanged -= Handler;
        }

        public string GetName() => "UnsubscribeHandler";
    }
}

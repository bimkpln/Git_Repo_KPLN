using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge;
using KPLN_RevitMcpBridge.ExternalCommands;
using KPLN_RevitMcpBridge.Server;

internal static class Program
{
    private static int _checks;
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
        _checks++;
    }
    private static Result Click(params TaskDialogResult[] choices)
    {
        TaskDialog.Results.Clear();
        TaskDialog.Instructions.Clear();
        TaskDialog.Actions.Clear();
        foreach (var choice in choices)
            TaskDialog.Results.Enqueue(choice);

        string message = null;
        return new BridgeStatusExtCmd().Execute(new ExternalCommandData(), ref message, new ElementSet());
    }

    [STAThread]
    private static void Main()
    {
        var app = new UIControlledApplication();
        var module = new Module();
        Check(module.Execute(app, "KPLN") == Result.Succeeded, "Loader registration failed");
        Check(module.Server == null && HttpBridgeServer.Created == 0, "Loader must not create a server");
        app.Pulse(); app.Pulse();
        Check(app.GetRibbonPanels("KPLN")[0].GetItems().Count == 1, "Register the ribbon once");
        Check(module.Server == null && HttpBridgeServer.Created == 0, "Idling must not start the server");
        Check(app.IdlingHandlers == 0, "Ribbon initialization should unsubscribe");
        var button = app.GetRibbonPanels("KPLN")[0].GetItems()[0];
        var images = KPLN_Loader.Application.ImageRequests;
        var assemblyName = typeof(Module).Assembly.GetName().Name;
        Check(images.Count == 2 && images[0] == (assemblyName, "mcpBridge", 16)
            && images[1] == (assemblyName, "mcpBridge", 32)
            && button.Image != null && button.LargeImage != null, "Loader must supply both icon sizes");
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
        var themedButtons = KPLN_Loader.Application.KPLNButtonsForImageReverse;
        Check(themedButtons.Count == 1 && themedButtons[0] == (button, "mcpBridge", assemblyName),
            "Register the actual button for Loader theme changes once");
#endif

        // Просмотр и закрытие окна при остановленном мосте
        Check(Click() == Result.Succeeded && HttpBridgeServer.Created == 0, "First click only shows status");
        Check(TaskDialog.LastInstruction == "Мост остановлен" && TaskDialog.Actions[0] == "Запустить мост",
            "Stopped bridge offers only Start");
        Click();
        Check(module.Server == null && HttpBridgeServer.Created == 0, "Repeated status views must not connect");


        // Явный запуск завершается без повторного открытия окна
        Check(Click(TaskDialogResult.CommandLink1) == Result.Succeeded && HttpBridgeServer.Live == 1,
            "Explicit Start connects");
        Check(TaskDialog.Instructions.Count == 1, "Start must not reopen the status dialog");
        var first = module.Server;
        Click();
        Check(ReferenceEquals(first, module.Server), "Viewing status must not restart the session");
        Check(TaskDialog.LastInstruction == "Мост работает" && TaskDialog.Actions[0] == "Остановить мост",
            "Reopening a running bridge offers only Stop");


        // Остановка и повторный просмотр не создают новую сессию
        Click(TaskDialogResult.CommandLink1);
        Check(first.Disposed && module.Server == null && HttpBridgeServer.Live == 0, "Stop closes the server");
        Check(TaskDialog.Instructions.Count == 1, "Stop must not reopen the status dialog");
        app.Pulse();
        Check(module.Server == null && HttpBridgeServer.Live == 0, "Stopped state persists through Idling");
        var createdBeforeView = HttpBridgeServer.Created;
        Click();
        Check(module.Server == null && HttpBridgeServer.Created == createdBeforeView,
            "Reopening after Stop must not start the bridge");
        Check(TaskDialog.LastInstruction == "Мост остановлен" && TaskDialog.Actions[0] == "Запустить мост",
            "Reopening a stopped bridge offers only Start");
        Click(TaskDialogResult.CommandLink1);
        Check(module.Server != null && HttpBridgeServer.Live == 1 && module.Server.SessionId != first.SessionId,
            "Only explicit Start creates a new session after Stop");
        module.Close();
        Check(Module.Current == null && HttpBridgeServer.Live == 0, "Loader shutdown must disconnect");

        var nextApp = new UIControlledApplication();
        var next = new Module(); next.Execute(nextApp, "KPLN");
        Check(next.Server == null && HttpBridgeServer.Live == 0, "Next load starts off despite previous connection");
        next.Close(); nextApp.Pulse();
        Check(nextApp.GetRibbonPanels("KPLN").Count == 0, "Close before first Idling cancels initialization");

        var retry = new Module(); retry.Execute(new UIControlledApplication(), "KPLN");
        HttpBridgeServer.FailStart = true;
        Check(Click() == Result.Succeeded && retry.Server == null, "Viewing status does not attempt a failing start");
        Check(Click(TaskDialogResult.CommandLink1) == Result.Cancelled && retry.Server == null && HttpBridgeServer.Live == 0,
            "Failed explicit Start leaves bridge off");
        Check(TaskDialog.LastError == "Simulated listener startup failure", "Show the startup error");
        HttpBridgeServer.FailStart = false;
        Click();
        Check(retry.Server == null, "Viewing status after failure must not retry");
        Check(Click(TaskDialogResult.CommandLink1) == Result.Succeeded && retry.LastError == null,
            "User can explicitly retry after failure");
        retry.Close();
        Console.WriteLine("Lifecycle: " + _checks + " checks passed.");
    }
}

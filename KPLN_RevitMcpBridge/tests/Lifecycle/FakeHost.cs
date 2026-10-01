using System;
using System.Collections.Generic;
using System.Windows.Media;

namespace Autodesk.Revit.Attributes
{
    public enum TransactionMode { Manual }
    public enum RegenerationOption { Manual }
    public class TransactionAttribute : Attribute { public TransactionAttribute(TransactionMode mode) { } }
    public class RegenerationAttribute : Attribute { public RegenerationAttribute(RegenerationOption option) { } }
}
namespace Autodesk.Revit.DB { public class ElementSet { } }
namespace Autodesk.Revit.UI.Events { public class IdlingEventArgs : EventArgs { } }
namespace Autodesk.Revit.UI
{
    public enum Result { Succeeded, Cancelled, Failed }
    public interface IExternalCommand { Result Execute(ExternalCommandData data, ref string message, DB.ElementSet elements); }
    public class RevitApplication { public string VersionNumber => "2023"; }
    public class UIApplication { public RevitApplication Application = new RevitApplication(); }
    public class ExternalCommandData { public UIApplication Application = new UIApplication(); }
    public class UIControlledApplication
    {
        private readonly List<RibbonPanel> _panels = new List<RibbonPanel>();
        public event EventHandler<Events.IdlingEventArgs> Idling;
        public int IdlingHandlers => Idling?.GetInvocationList().Length ?? 0;
        public void Pulse() { Idling?.Invoke(this, new Events.IdlingEventArgs()); }
        public List<RibbonPanel> GetRibbonPanels(string tabName) => _panels;
        public RibbonPanel CreateRibbonPanel(string tabName, string name)
        {
            var panel = new RibbonPanel { Name = name }; _panels.Add(panel); return panel;
        }
    }
    public class RibbonPanel
    {
        public string Name { get; set; }
        private readonly List<PushButton> _items = new List<PushButton>();
        public List<PushButton> GetItems() => _items;
        public PushButton AddItem(PushButtonData data)
        {
            var button = new PushButton(data.Name); _items.Add(button); return button;
        }
    }
    public class PushButtonData
    {
        public string Name { get; }
        public string ToolTip { get; set; }
        public string LongDescription { get; set; }
        public ImageSource Image { get; set; }
        public ImageSource LargeImage { get; set; }
        public PushButtonData(string name, string label, string assembly, string command) { Name = name; }
    }
    public class PushButton : PushButtonData
    {
        public string ItemText { get; set; }
        public PushButton(string name) : base(name, null, null, null) { }
    }
    public enum TaskDialogCommonButtons { Close }
    public enum TaskDialogCommandLinkId { CommandLink1, CommandLink2 }
    public enum TaskDialogResult { Close, CommandLink1, CommandLink2 }
    public class TaskDialog
    {
        public static readonly Queue<TaskDialogResult> Results = new Queue<TaskDialogResult>();
        public static readonly List<string> Instructions = new List<string>();
        public static readonly List<string> Actions = new List<string>();
        public static string LastError;
        public static string LastInstruction;
        private readonly List<string> _actions = new List<string>();
        public string MainInstruction { get; set; }
        public string MainContent { get; set; }
        public TaskDialogCommonButtons CommonButtons { get; set; }
        public TaskDialogResult DefaultButton { get; set; }
        public TaskDialog(string title) { }
        public void AddCommandLink(TaskDialogCommandLinkId id, string text) { _actions.Add(text); }
        public TaskDialogResult Show()
        {
            LastInstruction = MainInstruction;
            Instructions.Add(MainInstruction);
            Actions.Add(string.Join("|", _actions));
            if (DefaultButton != TaskDialogResult.Close)
                throw new Exception("Viewing status must default to Close");

            return Results.Count == 0 ? TaskDialogResult.Close : Results.Dequeue();
        }
        public static void Show(string title, string text) { LastError = text; }
    }
}
namespace KPLN_Loader
{
    public static class Application
    {
        public static readonly List<(Autodesk.Revit.UI.PushButton RButton, string IconBaseName, string KPLNPluginAssembleName)> KPLNButtonsForImageReverse
            = new List<(Autodesk.Revit.UI.PushButton, string, string)>();
        public static readonly List<(string Assembly, string Name, int Size)> ImageRequests
            = new List<(string, string, int)>();
        public static ImageSource GetBtnImage_ByTheme(string assemblyName, string iconBaseName, int size)
        {
            ImageRequests.Add((assemblyName, iconBaseName, size));
            return new DrawingImage();
        }
    }
}
namespace KPLN_Loader.Common
{
    public interface IExternalModule
    {
        Autodesk.Revit.UI.Result Execute(Autodesk.Revit.UI.UIControlledApplication app, string tabName);
        Autodesk.Revit.UI.Result Close();
    }
}
namespace KPLN_RevitMcpBridge.Server
{
    internal sealed class HttpBridgeServer : IDisposable
    {
        public static int Created;
        public static int Started;
        public static int Live;
        public static bool FailStart;
        private bool _running;
        public bool Disposed;
        public string Url => "http://127.0.0.1:18765";
        public string SessionId { get; } = Guid.NewGuid().ToString("N");
        public HttpBridgeServer(Autodesk.Revit.UI.UIControlledApplication app) { Created++; }
        public void Start()
        {
            if (FailStart) throw new InvalidOperationException("Simulated listener startup failure");
            Started++; Live++; _running = true;
        }
        public void Dispose() { if (_running) { Live--; _running = false; } Disposed = true; }
    }
}

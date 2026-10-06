using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_RevitMcpBridge.ExternalCommands;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Linq;
using System.Reflection;

namespace KPLN_RevitMcpBridge
{
    public sealed class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
        internal static Module Current { get; private set; }
        internal HttpBridgeServer Server { get; private set; }
        internal string LastError { get; private set; }
        private UIControlledApplication _application;
        private string _tabName;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            Current = this;
            _application = application;
            _tabName = tabName;
            // BIMTools creates BIM unconditionally. Wait until all Loader modules
            // have executed, so module ordering cannot create a duplicate panel.
            application.Idling += InitializeRibbon;


            // Загрузка регистрирует только кнопку. Сервер запускается отдельным
            // действием «Запустить мост» в окне состояния.
            return Result.Succeeded;
        }

        private void InitializeRibbon(object sender, IdlingEventArgs args)
        {
            _application.Idling -= InitializeRibbon;
            try
            {
                var panel = _application.GetRibbonPanels(_tabName).FirstOrDefault(p => p.Name == "BIM")
                    ?? _application.CreateRibbonPanel(_tabName, "BIM");
                if (panel.GetItems().Any(p => p.Name == "KPLN_RevitMcpBridge")) return;
                AddPushButtonDataInPanel(
                    "KPLN_RevitMcpBridge",
                    "Codex\nМост",
                    "Состояние моста Codex / управление подключением",
                    string.Format(
                        "Настройка моста для подключения к Codex.\n\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(BridgeStatusExtCmd).FullName,
                    panel,
                    "mcpBridge");
            }
            catch (Exception ex) { LastError = "Кнопка BIM: " + ex.Message; }
        }

        /// <summary>
        /// Добавляет кнопку по шаблону KPLN. imageName — базовое имя ресурса
        /// Imagens/{imageName}{16|32}[_dark].png; тему выбирает Loader.
        /// </summary>
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription,
            string longDescription, string className, RibbonPanel panel, string imageName)
        {
            var data = new PushButtonData(name, text, _assemblyPath, className);
            var button = (PushButton)panel.AddItem(data);
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Loader обновляет обе иконки при смене темы Revit 2024.
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, _assemblyName));
#endif
        }

        internal void Start()
        {
            Stop();
            try
            {
                Server = new HttpBridgeServer(_application);
                Server.Start(); LastError = null;
            }
            catch (Exception ex) { LastError = ex.Message; Stop(); }
        }
        internal void Stop() { Server?.Dispose(); Server = null; }
        public Result Close()
        {
            _application.Idling -= InitializeRibbon;
            Stop(); Current = null; return Result.Succeeded;
        }
    }
}

using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker;       
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Docking
{
    internal sealed class FamilyManagerPaneProvider : IDockablePaneProvider
    {
        private readonly FamilyManager _pane;
        public FamilyManagerPaneProvider(FamilyManager pane)
        {
            _pane = pane ?? throw new ArgumentNullException(nameof(pane));
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = _pane;

            var state = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser 
            };

            data.InitialState = state;
        }
    }

    internal static class FamilyManagerDock
    {
        public static readonly DockablePaneId PaneId =
            new DockablePaneId(new Guid("E5C4A9B3-6B41-4A0E-9F2C-7F7F2A2B9C11"));

        private static FamilyManager _paneInstance;

        public static void Register(UIControlledApplication app)
        {
            if (_paneInstance != null) return;

            var current = DBMainService.CurrentUserDBSubDepartment;
            string currentStr = current != null ? $"{current.Code}" : "Нет данных";

            _paneInstance = new FamilyManager(currentStr);

            var provider = new FamilyManagerPaneProvider(_paneInstance);
            app.RegisterDockablePane(PaneId, "KPLN. Менеджер семейств", provider);
        }

        public static void Toggle(UIApplication uiapp)
        {
            if (uiapp == null) throw new ArgumentNullException(nameof(uiapp));
            var pane = uiapp.GetDockablePane(PaneId);
            if (pane.IsShown())
                pane.Hide();
            else
                pane.Show();
        }
    }
}
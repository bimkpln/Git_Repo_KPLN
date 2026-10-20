using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_ViewsAndLists_CopyByLevels.ExternalCommands;
using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;

namespace KPLN_ViewsAndLists_CopyByLevels
{
    public sealed class Module : IExternalModule
    {
        private UIControlledApplication _application;
        private string _tabName;
        private int _attempts;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            _application = application;
            _tabName = tabName;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Панель создаёт KPLN_ViewsAndLists_Ribbon. Порядок загрузки модулей не фиксирован.
            if (!AttachButton())
                application.Idling += OnIdling;

            return Result.Succeeded;
        }

        private void OnIdling(object sender, IdlingEventArgs args)
        {
            if (AttachButton())
            {
                _application.Idling -= OnIdling;
            }
            else if (++_attempts >= 10)
            {
                _application.Idling -= OnIdling;
                TaskDialog.Show("KPLN: Листы по этажам", "Панель «Виды и листы» не найдена. Подключите модуль KPLN_ViewsAndLists_Ribbon и перезапустите Revit.");
            }
        }

        private bool AttachButton()
        {
            RibbonPanel panel = _application.GetRibbonPanels(_tabName).FirstOrDefault(p => p.Name == "Виды и листы");
            if (panel == null) return false;
            if (panel.GetItems().Any(i => i.Name == ModuleData.ModuleName)) return true;

            var data = new PushButtonData(ModuleData.ModuleName, "Листы\nпо этажам", Assembly.GetExecutingAssembly().Location,
                typeof(CopyByLevelsExtCmd).FullName)
            {
                ToolTip = "Создать листы по выбранным уровням на основе эталона.",
                LongDescription = "Планы, оформление листа, легенды и экспликации с выборочной заменой фильтров.\nВерсия: " + ModuleData.Version,
                Image = CreateIcon(),
                LargeImage = CreateIcon()
            };
            panel.AddItem(data);
            return true;
        }

        private static ImageSource CreateIcon()
        {
            var drawing = new DrawingGroup();
            using (DrawingContext dc = drawing.Open())
            {
                dc.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, 32, 32));
                var pen = new Pen(new SolidColorBrush(Color.FromRgb(63, 140, 255)), 1.8);
                dc.DrawRoundedRectangle(Brushes.WhiteSmoke, pen, new Rect(3, 3, 19, 23), 1, 1);
                dc.DrawRoundedRectangle(Brushes.WhiteSmoke, pen, new Rect(10, 8, 19, 23), 1, 1);
                dc.DrawLine(pen, new Point(14, 14), new Point(25, 14));
                dc.DrawLine(pen, new Point(14, 19), new Point(25, 19));
                dc.DrawLine(pen, new Point(14, 25), new Point(25, 25));
            }
            var image = new DrawingImage(drawing);
            image.Freeze();
            return image;
        }

        public Result Close()
        {
            if (_application != null) _application.Idling -= OnIdling;
            return Result.Succeeded;
        }
    }
}

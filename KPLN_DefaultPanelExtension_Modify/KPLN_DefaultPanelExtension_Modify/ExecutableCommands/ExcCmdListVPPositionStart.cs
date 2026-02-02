using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace KPLN_DefaultPanelExtension_Modify.ExecutableCommands
{
    internal class ExcCmdListVPPositionStart : IExecutableCommand
    {
        internal const string PluginName = "Положение вида";

        private readonly Element[] _selElems;

        public ExcCmdListVPPositionStart(Element[] selElems)
        {
            _selElems = selElems;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;


            Viewport[] selVPs = _selElems
                .Where(el => el is Viewport)
                .Cast<Viewport>()
                .ToArray();

            if (selVPs.Length == 0)
            {
                MessageBox.Show("Ошибка - в выборку не попали видовые экраны. Работа остановлена", "Внимание", MessageBoxButton.OK, MessageBoxImage.Error);
                return Result.Cancelled;
            }


            //ListVPPositionMainFrom mainFrom = new ListVPPositionMainFrom(_doc, selVPs);
            //WindowHandleSearch.MainWindowHandle.SetAsOwner(mainFrom);
            //mainFrom.Show();

            return Result.Succeeded;
        }
    }
}

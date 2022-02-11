using Autodesk.Revit.UI;

namespace KPLN_Loader.Common
{
    /// <summary>
    /// Интерфейс для запуска команд Revit (событие OnIdling запускает команду внутри текущего UIControlledApplication)
    /// </summary>
    public interface IExecutableCommand
    {
        Result Execute(UIApplication app);
    }
}

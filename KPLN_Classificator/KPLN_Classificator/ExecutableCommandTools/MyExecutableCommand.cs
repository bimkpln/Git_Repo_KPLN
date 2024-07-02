using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_Classificator
{
    /// <summary>
    /// Интерфейс для запуска команд Revit (событие OnIdling запускает команду внутри текущего UIControlledApplication)
    /// </summary>
    public interface MyExecutableCommand : IExecutableCommand
    {
        Result MyExecute(UIApplication app);
    }
}

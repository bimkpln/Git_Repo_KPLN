using Autodesk.Revit.UI;
#if Revit2020 || Revit2018
using KPLN_Loader.Common;
#endif

namespace KPLN_Classificator
{
    /// <summary>
    /// Интерфейс для запуска команд Revit (событие OnIdling запускает команду внутри текущего UIControlledApplication)
    /// </summary>
    public interface MyExecutableCommand
#if Revit2020 || Revit2018
        : IExecutableCommand
#endif
    {
        Result Execute(UIApplication app);
    }
}

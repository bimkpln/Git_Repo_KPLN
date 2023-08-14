using Autodesk.Revit.UI;

namespace KPLN_Loader.Common
{
    /// <summary>
    /// Интерфейс для создания подгружаемых модулей текущего плагина KPLN_Loader. Метод Execute запускается вместе с Revit (OnStartUp), а метод Close - при закрытии (OnShutDown)
    /// </summary>
    public interface IExternalModule
    {
        Result Execute(UIControlledApplication application, string tabName);
        
        Result Close();
    }
}

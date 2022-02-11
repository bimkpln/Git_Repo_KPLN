using Autodesk.Revit.UI;

namespace KPLN_Loader.Common
{
    public interface IExternalModule
    {
        /// <summary>
        /// Интерфейс для создания подгружаемых модулей текущего плагина KPLN_Loader. Метод Execute запускается вместе с Revit (OnStartUp), а метод Close - при закрытии (OnShutDown)
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        Result Execute(UIControlledApplication application, string tabName);
        Result Close();
    }
    
}

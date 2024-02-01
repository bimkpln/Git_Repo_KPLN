using Autodesk.Revit.UI;

namespace KPLN_BIMTools_Ribbon.Common
{
    internal interface IExcecuteByUIApp
    {
        /// <summary>
        /// Спец. метод для вызова данного класса внутри плагина: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        Result ExecuteByUIApp(UIApplication uiapp);
    }
}

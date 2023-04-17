using Autodesk.Revit.UI;
using System.Reflection;

namespace KPLN_Tools.Common
{
    /// <summary>
    /// Extension UI.RibbonItem для поиска ID элемента UI.RibbonItem
    /// </summary>
    internal static class RibbonItemExtensions
    {
        public static string GetId(this RibbonItem ribbonItem)
        {
            var type = typeof(RibbonItem);

            var parentId = type
                .GetField("m_parentId", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.GetValue(ribbonItem) ?? string.Empty;

            var generateIdMethod = type
                .GetMethod("generateId", BindingFlags.Static | BindingFlags.NonPublic);

            return (string)generateIdMethod?.Invoke(ribbonItem, new[] { parentId, ribbonItem.Name });
        }
    }
}

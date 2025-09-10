using HtmlAgilityPack;
using System.Linq;
using System.Text;
using System.Web;

namespace KPLN_Clashes_Ribbon.Tools
{
    /// <summary>
    /// Helper methods for parsing NavisWorks HTML reports.
    /// </summary>
    internal static class HtmlReportParser
    {
        public static string GetId(HtmlNode node, int num)
        {
            var element = node.ChildNodes
                .FirstOrDefault(n => Decode(n.GetAttributeValue("class", string.Empty)) == $"элемент{num}Содержимое");
            if (element == null)
                return null;

            string text = Decode(element.InnerText);
            return text.ToLower().Contains("id объекта")
                ? text.Split(':').LastOrDefault()?.Trim()
                : null;
        }

        public static string GetFullName(HtmlNode node, int num)
        {
            var element = node.ChildNodes
                .FirstOrDefault(n => Decode(n.GetAttributeValue("class", string.Empty)) == $"элемент{num}Содержимое");
            if (element == null)
                return null;

            string text = Decode(element.InnerText);
            string lower = text.ToLower();
            if (lower.Contains(".rvt") || lower.Contains(".nwc") || lower.Contains(".nwd"))
            {
                return text.Trim();
            }
            return null;
        }

        public static string GetImage(HtmlNode node)
        {
            var linkNode = node.Descendants("a").FirstOrDefault();
            return linkNode == null
                ? null
                : HttpUtility.UrlDecode(linkNode.GetAttributeValue("href", string.Empty));
        }

        public static string GetPoint(HtmlNode node)
        {
            return node.ChildNodes
                .Select(n => Decode(n.InnerText))
                .FirstOrDefault(t => t.ToLower().Contains("x:") && t.ToLower().Contains("y:") && t.ToLower().Contains("z:"))
                ?.Trim();
        }

        public static string GetHeader(HtmlNode node)
        {
            return node.ChildNodes
                .Select(n => Decode(n.InnerText))
                .FirstOrDefault(t => t.ToLower().StartsWith("конфликт") && t.ToLower() != "конфликты")
                ?.Trim();
        }

        public static string Optimize(string value) => value?.Trim() ?? string.Empty;

        public static string Decode(string value)
        {
            byte[] bytes = Encoding.Default.GetBytes(value);
            return Encoding.UTF8.GetString(bytes);
        }
    }
}

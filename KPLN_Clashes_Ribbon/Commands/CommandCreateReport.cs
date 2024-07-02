using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCreateReport : IExternalCommand
    {
        private static string Decode(string value)
        {
            string myString = value;
            byte[] bytes = Encoding.Default.GetBytes(myString);
            myString = Encoding.UTF8.GetString(bytes);
            return myString;
        }
        public string GetHeader(HtmlNode node)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (Decode(sub_node.InnerText).ToLower().StartsWith("конфликт") && Decode(sub_node.InnerText).ToLower() != "конфликты")
                {
                    return Decode(sub_node.InnerText);
                }
            }
            return null;
        }
        public string GetPoint(HtmlNode node)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (Decode(sub_node.InnerText).ToLower().Contains("x:") && Decode(sub_node.InnerText).ToLower().Contains("y:") && Decode(sub_node.InnerText).ToLower().Contains("z:"))
                {
                    return Decode(sub_node.InnerText);
                }
            }
            return null;
        }
        public string GetId(HtmlNode node, int num)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (Decode(sub_node.GetAttributeValue("class", "NONE")) == string.Format("элемент{0}Содержимое", num.ToString()))
                {
                    if (Decode(sub_node.InnerText).ToLower().Contains("id объекта"))
                    {
                        return Decode(sub_node.InnerText).Split(':').Last();
                    }
                }
            }
            return null;
        }
        public string GetFullName(HtmlNode node, int num)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (Decode(sub_node.GetAttributeValue("class", "NONE")) == string.Format("элемент{0}Содержимое", num.ToString()))
                {
                    if (Decode(sub_node.InnerText).ToLower().Contains(".rvt") || Decode(sub_node.InnerText).ToLower().Contains(".nwc") || Decode(sub_node.InnerText).ToLower().Contains(".nwd"))
                    {
                        return Decode(sub_node.InnerText);
                    }
                }
            }
            return null;
        }
        public string GetImage(HtmlNode node)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                foreach (HtmlNode sub_sub_node in sub_node.ChildNodes)
                {
                    if (sub_sub_node.Name == "a")
                    {
                        return HttpUtility.UrlDecode(sub_sub_node.GetAttributeValue("href", "NONE"));
                    }
                }
            }
            return null;
        }
        public string Optimize(string value)
        {
            List<char> final_chars = new List<char>();
            List<char> chars = new List<char>();
            foreach (char c in value)
            {
                if (char.IsWhiteSpace(c) && chars.Count == 0)
                {
                    continue;
                }
                chars.Add(c);
            }
            chars.Reverse();
            foreach (char c in chars)
            {
                if (char.IsWhiteSpace(c) && final_chars.Count == 0)
                {
                    continue;
                }
                final_chars.Add(c);
            }
            final_chars.Reverse();
            string result = string.Empty;
            foreach (char c in final_chars)
            {
                result += c;
            }
            return result;
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog
                {
                    Filter = "html report (*.html)|*.html",
                    Title = "Выберите отчет NavisWorks в формате .html",
                    CheckFileExists = true,
                    CheckPathExists = true,
                    Multiselect = false
                };
                DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    FileInfo file = new FileInfo(dialog.FileName);
                    Print(file.Extension, MessageType.System_Regular);
                    if (file.Extension == ".html")
                    {
                        HtmlAgilityPack.HtmlDocument htmlSnippet = new HtmlAgilityPack.HtmlDocument();
                        using (FileStream stream = file.OpenRead())
                        {
                            htmlSnippet.Load(stream);
                        }
                        foreach (HtmlNode link in htmlSnippet.DocumentNode.SelectNodes("//tr"))
                        {
                            string header = GetHeader(link);
                            if (header != null)
                            {
                                FileInfo imageFile = new FileInfo(Path.Combine(file.DirectoryName, Optimize(GetImage(link))));
                                if (imageFile.Exists)
                                {
                                    Stream stream = System.IO.File.Open(imageFile.FullName, FileMode.Open);
                                }
                                Print(string.Format("{0}", header), MessageType.Success);
                                Print(string.Format("Image = '{0}'", Optimize(GetImage(link))), MessageType.System_Regular);
                                Print(string.Format("Point = '{0}'", Optimize(GetPoint(link))), MessageType.System_Regular);
                                Print(string.Format("ID1 = '{0}'", Optimize(GetId(link, 1))), MessageType.System_Regular);
                                Print(string.Format("ID2 = '{0}'", Optimize(GetId(link, 2))), MessageType.System_Regular);
                                Print(string.Format("Name1 = '{0}'", GetFullName(link, 1)), MessageType.System_Regular);
                                Print(string.Format("Name2 = '{0}'", GetFullName(link, 2)), MessageType.System_Regular);
                            }
                        }
                    }
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
    }
}

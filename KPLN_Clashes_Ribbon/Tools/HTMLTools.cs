using HtmlAgilityPack;
using KPLN_Clashes_Ribbon.Core.Reports;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Web;

namespace KPLN_Clashes_Ribbon.Tools
{
    internal static class HTMLTools
    {
        public static string GetTime(string value)
        {
            string result = string.Empty;
            int i = 0;
            foreach (char c in value)
            {
                if (i >= 5)
                {
                    break;
                }
                result += c;
                i++;
            }
            return result;
        }
        public static string GetMessage(string value)
        {
            string result = string.Empty;
            int i = 0;
            foreach (char c in value)
            {
                if (i >= 5)
                {
                    result += c;
                }
                i++;
            }
            return result;
        }
        public static string TryGetComments(string row)
        {
            string comment = string.Empty;
            string[] lines = row.Split('#');
            foreach (string line in lines)
            {
                if (line != string.Empty)
                {
                    List<string> parts = new List<string>();
                    string[] pp = line.Split(' ');
                    foreach (string p in pp)
                    {
                        if (p != string.Empty) { parts.Add(p); }
                    }
                    string user = parts[2];
                    string time = parts[4] + " " + GetTime(parts[5]) + ":00";
                    string message = GetMessage(parts[5]);
                    List<string> msgParts = new List<string>() { message };
                    int i = 0;
                    foreach (string p in parts)
                    {
                        if (i > 5)
                        {
                            msgParts.Add(p);
                        }
                        i++;
                    }
                    comment = string.Join(" ", msgParts);
                }
            }

            return comment;
        }
        public static int GetRowId(List<string> values, string value, bool skipfirst = false)
        {
            bool first_found = false;
            int n = 0;
            foreach (string v in values)
            {
                if (v == value)
                {
                    if (!skipfirst)
                    {
                        return n;
                    }
                    else
                    {
                        if (first_found)
                        {
                            return n;
                        }
                        else
                        {
                            first_found = true;
                        }
                    }
                }
                n++;
            }
            return -1;
        }
        public static string GetValue(HtmlNode node, int element, bool decode)
        {
            try
            {
                if (decode)
                {
                    return Decode(node.ChildNodes[element].InnerText);
                }
                else
                {
                    return node.ChildNodes[element].InnerText;
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
        public static List<string> GetHeaders(HtmlNode node, bool decode)
        {
            List<string> headers = new List<string>();
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (decode)
                {
                    headers.Add(Decode(sub_node.InnerText));
                }
                else
                {
                    headers.Add(sub_node.InnerText);
                }

            }
            return headers;
        }
        public static bool IsMainHeader(HtmlNode node, out List<string> result, out bool decode)
        {
            List<string> headers = new List<string>();
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                string value = Decode(sub_node.InnerText);
                if (value.Length > 0)
                {
                    headers.Add(value);
                }
            }
            if (headers.Contains("Наименование конфликта") && headers.Contains("Изображение") && headers.Contains("Объект Id") && headers.Contains("Путь"))
            {
                result = headers;
                decode = true;
                return true;
            }
            else
            {
                headers.Clear();
            }
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                string value = sub_node.InnerText;
                if (value.Length > 0)
                {
                    headers.Add(value);
                }
            }
            if (headers.Contains("Наименование конфликта") && headers.Contains("Изображение") && headers.Contains("Объект Id") && headers.Contains("Путь"))
            {
                result = headers;
                decode = false;
                return true;
            }
            result = new List<string>();
            decode = false;
            return false;
        }
        public static string GetImage(HtmlNode node, int num)
        {
            foreach (HtmlNode sub_sub_node in node.ChildNodes[num].ChildNodes)
            {
                if (sub_sub_node.Name == "a")
                {
                    return HttpUtility.UrlDecode(sub_sub_node.GetAttributeValue("href", "NONE"));
                }
            }
            return null;
        }
        public static string GetId(string value)
        {
            return value.Split(':').Last();
        }
        //
        private static string Decode(string value)
        {
            string myString = value;
            byte[] bytes = Encoding.Default.GetBytes(myString);
            myString = Encoding.UTF8.GetString(bytes);
            return myString;
        }

        public static string GetFullName(HtmlNode node, int num)
        {
            foreach (HtmlNode sub_node in node.ChildNodes)
            {
                if (Decode(sub_node.GetAttributeValue("class", "NONE")) == string.Format("элемент{0}Содержимое", num.ToString()))
                {
                    if (Decode(sub_node.InnerText).ToLower().Contains(".rvt") || Decode(sub_node.InnerText).ToLower().Contains(".nwc") || Decode(sub_node.InnerText).ToLower().Contains(".nwd"))
                    {
                        return OptimizeV(Decode(sub_node.InnerText));
                    }
                }
            }
            return null;
        }
        public static string Optimize(string value)
        {
            try
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
            catch (Exception)
            {
                return "NONE";
            }

        }
        public static string OptimizeV(string value)
        {
            try
            {
                string optimized = value.Replace("&gt;", "➜");
                List<char> chars = new List<char>();
                bool last_space = false;
                foreach (char c in optimized)
                {
                    if (char.IsWhiteSpace(c) && !last_space)
                    {
                        chars.Add(' ');
                        last_space = true;
                    }
                    else
                    {
                        if (char.IsDigit(c) || char.IsLetter(c) || "➜_-().,[]';?!@#$%*&".Contains(c))
                        {
                            chars.Add(c);
                            last_space = false;
                        }
                    }


                }
                string input = string.Empty;
                foreach (char c in chars)
                {
                    input += c;
                }
                return input;
            }
            catch (Exception)
            {
                return "NONE";
            }

        }
    }
}

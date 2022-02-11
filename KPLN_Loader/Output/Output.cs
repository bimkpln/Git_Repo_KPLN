using HtmlAgilityPack;

using KPLN_Loader.Forms;
using static KPLN_Loader.Preferences;

using System;
using System.IO;
using System.Collections.Generic;

namespace KPLN_Loader.Output
{
    public static class Output
    {
        public static Queue<string> LogQueue = new Queue<string>();
        public static OutputWindow FormOutput { get; set; }
        public static HtmlDocument htmlDocument = new HtmlDocument();
        public static string outputPath = string.Format(@"{0}\log_{1}.html", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) , Guid.NewGuid());
        public static void Main()
        {
            CreateStream();
        }
        private static void OutputPrint(string value, string cssclass)
        {
            string insertablevalue = string.Join("<b>«", value.Split('«'));
            insertablevalue = string.Join("»</b>", insertablevalue.Split('»'));
            insertablevalue = string.Join("<b>[", insertablevalue.Split('['));
            insertablevalue = string.Join("]</b>", insertablevalue.Split(']'));
            if (FormOutput == null)
            {
                FormOutput = new OutputWindow();
                CreateStream();
                FormOutput.Show();
                FormOutput.webBrowser.Navigate(new Uri(outputPath));
            }
            else
            {
                FormOutput.Show();
            }
            HtmlNode body = htmlDocument.DocumentNode.SelectSingleNode("/html/body");
            if (body == null)
            {
                CreateStream();
                body = htmlDocument.DocumentNode.SelectSingleNode("/html/body");
            }
            HtmlNode node = HtmlNode.CreateNode(string.Format("<p class='{0}'>{1}</p>", cssclass, insertablevalue));
            body.InsertBefore(node, body.FirstChild);
            using (StreamWriter streamWriter = new StreamWriter(outputPath))
            {
                htmlDocument.Save(streamWriter);
                streamWriter.Close();
            }
            FormOutput.webBrowser.Refresh();
            System.Windows.Forms.Application.DoEvents();
            FormOutput.webBrowser.Refresh();
            FormOutput.BringToFront();
        }
        public static void PrintError(Exception e)
        {
            try
            {
                OutputPrint(e.StackTrace, "code");
                OutputPrint(e.Message, "logerrorheader");
            }
            catch (Exception) { }
        }
        public static void PrintError(Exception e, string message)
        {
            try
            {
                OutputPrint(e.StackTrace, "code");
                OutputPrint(string.Format("{0}<br>   {1}", message, e.Message), "logerrorheader");
            }
            catch (Exception) { }
        }
        public static void Print(string value, MessageType type)
        {
            try
            {
                switch (type)
                {
                    case MessageType.Error:
                        OutputPrint(value, "logerror");
                        break;
                    case MessageType.Header:
                        OutputPrint(value, "logheader");
                        break;
                    case MessageType.Success:
                        OutputPrint(value, "logsuccess");
                        break;
                    case MessageType.Warning:
                        OutputPrint(value, "logwarning");
                        break;
                    case MessageType.Critical:
                        OutputPrint(value, "logcritical");
                        break;
                    case MessageType.Code:
                        OutputPrint(value, "code");
                        break;
                    case MessageType.System_OK:
                        OutputPrint(value, "systemok");
                        break;
                    case MessageType.System_Regular:
                        OutputPrint(value, "systemregular");
                        break;
                    default:
                        OutputPrint(value, "logdefault");
                        break;
                }
            }
            catch (Exception) { }
        }
        private static void CreateStream()
        {
            htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(string.Format(@"<html>{0}{1}<body></body></html>", HTML_Output_DocType, HTML_Output_Head));
            //
            using (StreamWriter streamWriter = new StreamWriter(outputPath))
            {
                htmlDocument.Save(streamWriter);
                streamWriter.Close();
            }
        }
    }
}

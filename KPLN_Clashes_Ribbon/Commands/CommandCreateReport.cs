using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HtmlAgilityPack;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.IO;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCreateReport : IExternalCommand
    {
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
                            string header = HtmlReportParser.GetHeader(link);
                            if (header != null)
                            {
                                FileInfo imageFile = new FileInfo(Path.Combine(file.DirectoryName, HtmlReportParser.Optimize(HtmlReportParser.GetImage(link))));
                                if (imageFile.Exists)
                                {
                                    Stream stream = System.IO.File.Open(imageFile.FullName, FileMode.Open);
                                }
                                Print(string.Format("{0}", header), MessageType.Success);
                                Print(string.Format("Image = '{0}'", HtmlReportParser.Optimize(HtmlReportParser.GetImage(link))), MessageType.System_Regular);
                                Print(string.Format("Point = '{0}'", HtmlReportParser.Optimize(HtmlReportParser.GetPoint(link))), MessageType.System_Regular);
                                Print(string.Format("ID1 = '{0}'", HtmlReportParser.Optimize(HtmlReportParser.GetId(link, 1))), MessageType.System_Regular);
                                Print(string.Format("ID2 = '{0}'", HtmlReportParser.Optimize(HtmlReportParser.GetId(link, 2))), MessageType.System_Regular);
                                Print(string.Format("Name1 = '{0}'", HtmlReportParser.GetFullName(link, 1)), MessageType.System_Regular);
                                Print(string.Format("Name2 = '{0}'", HtmlReportParser.GetFullName(link, 2)), MessageType.System_Regular);
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
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Cancelled;
            }
        }
    }
}

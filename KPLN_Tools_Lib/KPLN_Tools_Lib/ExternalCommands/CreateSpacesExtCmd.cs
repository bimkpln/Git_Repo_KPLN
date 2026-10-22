using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools_Lib.Forms;
using KPLN_Tools_Lib.Forms.Entities;
using KPLN_Tools_Lib.Services;
using System;
using System.Linq;
using System.Net;

namespace KPLN_Tools_Lib.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CreateSpacesExtCmd : IExternalCommand
    {
        public const string PluginName = "Создать пространства";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document document = commandData.Application.ActiveUIDocument?.Document;
                if (document == null || document.IsFamilyDocument || document.IsReadOnly)
                {
                    TaskDialog.Show(PluginName, "Откройте проект, доступный для редактирования.");
                    return Result.Cancelled;
                }

                RevitLinkInstance[] links = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                    .Where(link => link.GetLinkDocument() != null).ToArray();

                if (links.Length == 0)
                {
                    TaskDialog.Show(PluginName, "В проекте нет загруженных связей Revit.");
                    return Result.Cancelled;
                }

                if (!new FilteredElementCollector(document).OfClass(typeof(Level)).Any())
                {
                    TaskDialog.Show(PluginName, "В текущем проекте нет уровней для создания пространств.");
                    return Result.Cancelled;
                }


                var model = new CreateSpacesM(document, links);
                var window = new CreateSpaces(model);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(window);

                if (window.ShowDialog() != true)
                    return Result.Cancelled;

                // Модальный диалог закрыт; команда продолжает выполняться в контексте Revit API.
                CreateSpacesResult result = new CreateSpacesService().Create(document,
                    model.Links.Where(link => link.IsSelected).Select(link => link.RevitLink), model.SelectedPhase);

                bool hasIssues = result.Messages.Count > 0 || result.FailureIdsByMessage.Count > 0;
                if (hasIssues)
                {
                    HtmlOutput.Print(PluginName, MessageType.Header);
                    if (result.Messages.Count > 0)
                    {
                        HtmlOutput.Print(string.Join("<br/>", result.Messages.Select(WebUtility.HtmlEncode)), MessageType.Warning);
                    }

                    HtmlOutput.PrintMsgDict("Предупреждение Revit", MessageType.Warning,
                        result.FailureIdsByMessage.ToDictionary(pair => WebUtility.HtmlEncode(pair.Key), pair => pair.Value));
                }

                TaskDialog.Show(PluginName, result.Summary + (hasIssues
                    ? "\nПодробности приведены в окне вывода KPLN." : string.Empty));
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }
    }
}

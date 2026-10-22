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
    public sealed class TagSpacesExtCmd : IExternalCommand
    {
        public const string PluginName = "Маркировать помещения/пространства";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDocument = commandData.Application.ActiveUIDocument;
                Document document = uiDocument?.Document;
                if (document == null || document.IsFamilyDocument || document.IsReadOnly)
                {
                    TaskDialog.Show(PluginName, "Откройте проект, доступный для редактирования.");
                    return Result.Cancelled;
                }

                Element[] selection = uiDocument.Selection.GetElementIds().Select(document.GetElement).ToArray();
                if (selection.Length == 0 || selection.Any(element => !IsSupportedView(element)))
                {
                    TaskDialog.Show(PluginName, "Выделите в диспетчере проекта один или несколько планов этажей / потолков "
                        + "через Ctrl или Shift. В выборе должны быть только планы, без шаблонов видов.");
                    return Result.Cancelled;
                }


                var model = new TagSpacesM(document, selection.Cast<ViewPlan>().OrderBy(view => view.Name).ToArray());
                model.ApplyConfig(TagSpacesConfigService.Load(document, out string configWarning));
                model.ConfigStatus = configWarning;
                
                var window = new TagSpaces(model);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(window);
                if (window.ShowDialog() != true)
                    return Result.Cancelled;

                // После модального окна остаёмся в исходном контексте Revit API.
                string saveWarning = TagSpacesConfigService.Save(document, model.GetConfig());
                var equipmentOptions = new TagSpacesEquipmentOptions(
                    model.IsSpace && model.UseRadiators ? model.SelectedRadiatorType.RevitType as Autodesk.Revit.DB.Mechanical.SpaceTagType : null,
                    model.IsSpace && model.UseGrilles ? model.SelectedGrilleType.RevitType as Autodesk.Revit.DB.Mechanical.SpaceTagType : null);
                TagSpacesResult result = new TagSpacesService().Create(document, model.Views,
                    model.SelectedType.RevitType, model.IsRoom ? model.Sources.Where(source => source.IsSelected).Select(source => source.RevitLink) : new RevitLinkInstance[] { null }, equipmentOptions);

                if (!string.IsNullOrEmpty(saveWarning))
                    result.Messages.Add(saveWarning);

                bool hasIssues = result.Messages.Count > 0 || result.Errors.Count > 0;
                if (hasIssues)
                {
                    HtmlOutput.Print(PluginName, MessageType.Header);
                    if (result.Messages.Count > 0)
                        HtmlOutput.Print(string.Join("<br/>", result.Messages.Select(WebUtility.HtmlEncode)), MessageType.Warning);
                    if (result.Errors.Count > 0)
                        HtmlOutput.Print(string.Join("<br/>", result.Errors.Select(WebUtility.HtmlEncode)), MessageType.Error);
                }

                TaskDialog.Show(PluginName, result.Summary + (hasIssues
                    ? "\nПодробности приведены в окне вывода KPLN." : string.Empty));
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }

        internal static bool IsSupportedView(Element element)
            => element is ViewPlan view && !view.IsTemplate && view.GenLevel != null
                && (view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan);
    }
}

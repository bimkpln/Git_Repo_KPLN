using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Finishing.CommandTools;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Finishing.Tools;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class SetColor : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            //
            ColorPresset.RandomR = new KPLNRandom(23452345 + (int)DateTime.Now.Ticks);
            ColorPresset.RandomG = new KPLNRandom(20904581 + (int)DateTime.Now.Ticks);
            ColorPresset.RandomB = new KPLNRandom(59827008 + (int)DateTime.Now.Ticks);
            //
            TaskDialog td = new TaskDialog("Окрасить по помещению");
            td.TitleAutoPrefix = false;
            td.MainContent = "Окрасить элементы отделки по помещениям?";
            td.FooterText = Names.task_dialog_hint;
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Close;
            TaskDialogResult result = td.Show();
            if (result != TaskDialogResult.Yes) { return Result.Cancelled; }
            try
            {
                View view = doc.ActiveView;
                if (view.ViewType == ViewType.ThreeD)
                {
                    using (Transaction t = new Transaction(doc, "KPLN Окрасить"))
                    {
                        t.Start();
                        List<ColorPresset> presets = new List<ColorPresset>();
                        FillPatternElement fillPatternElement = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().First(a => a.GetFillPattern().IsSolidFill);
                        OverrideGraphicSettings settingsDefault = new OverrideGraphicSettings();
                        Color white = new Color(255, 255, 255);
                        settingsDefault.SetSurfaceForegroundPatternColor(white);
                        settingsDefault.SetProjectionLineWeight(1);
                        settingsDefault.SetHalftone(true);
                        settingsDefault.SetSurfaceForegroundPatternId(fillPatternElement.Id);
                        settingsDefault.SetSurfaceTransparency(80);
                        //
                        OverrideGraphicSettings settingsError = new OverrideGraphicSettings();
                        Color red = new Color(255, 40, 0);
                        settingsError.SetSurfaceForegroundPatternColor(red);
                        settingsError.SetProjectionLineWeight(7);
                        settingsError.SetHalftone(false);
                        settingsError.SetSurfaceForegroundPatternId(fillPatternElement.Id);
                        settingsError.SetSurfaceTransparency(0);
#if Revit2018
                        settingsDefault.SetProjectionFillColor(white);
                        settingsDefault.SetProjectionLineWeight(1);
                        settingsDefault.SetHalftone(true);
                        settingsDefault.SetProjectionFillPatternId(fillPatternElement.Id);
                        settingsDefault.SetSurfaceTransparency(80);
                        //
                        OverrideGraphicSettings settingsError = new OverrideGraphicSettings();
                        Color red = new Color(255, 40, 0);
                        settingsError.SetProjectionFillColor(red);
                        settingsError.SetProjectionLineWeight(7);
                        settingsError.SetHalftone(false);
                        settingsError.SetProjectionFillPatternId(fillPatternElement.Id);
                        settingsError.SetSurfaceTransparency(0);
#endif
                        int n = new FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements().Count();
                        n += new FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements().Count();
                        n += new FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements().Count();
                        string s = "{0} из " + n.ToString() + " элементов обработано";
                        using (ProgressFormSimple pf = new ProgressFormSimple("Окрашивание элементов", s, n))
                        {
                            foreach (BuiltInCategory cat in new BuiltInCategory[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings })
                            {
                                foreach (Element element in new FilteredElementCollector(doc, view.Id).OfCategory(cat).WhereElementIsNotElementType().ToElements())
                                {
                                    pf.Increment();
                                    try
                                    {
                                        Element type = GetTypeElement(element);
                                        if (type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().ToLower() == "отделка")
                                        {
                                            try
                                            {
                                                int roomId = int.Parse(element.LookupParameter("О_Id помещения").AsString(), System.Globalization.NumberStyles.Integer);
                                                Room room = doc.GetElement(new ElementId(roomId)) as Room;
                                                ColorPresset preset = ColorPresset.GetRoomColor(presets, roomId, fillPatternElement);
                                                view.SetElementOverrides(element.Id, preset.Settings);
                                            }
                                            catch (Exception)
                                            {
                                                view.SetElementOverrides(element.Id, settingsError);
                                            }
                                        }
                                        else
                                        {
                                            view.SetElementOverrides(element.Id, settingsDefault);
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        try
                                        {
                                            view.SetElementOverrides(element.Id, settingsDefault);
                                        }
                                        catch (Exception) { }
                                    }
                                }
                            }
                        }
                        t.Commit();
                    }
                }
            }
            catch (Exception e)
            {
                HtmlOutput.PrintError(e);
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
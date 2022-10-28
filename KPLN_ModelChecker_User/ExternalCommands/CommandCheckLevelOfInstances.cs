using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_ModelChecker_User.Common.Collections;
using static KPLN_Loader.Output.Output;
using System.Collections.ObjectModel;
using KPLN_ModelChecker_User.Forms;
using static KPLN_ModelChecker_User.Tools.Extentions;
namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckLevelOfInstances : IExternalCommand
    {
        public static BuiltInCategory[] CategoriesToCheck = new BuiltInCategory[] {
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_GenericModel };

        private BuiltInParameter[] parameters = new BuiltInParameter[] { 
            BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM, 
            BuiltInParameter.WALL_BASE_OFFSET,
            BuiltInParameter.WALL_TOP_OFFSET,
            BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,
            BuiltInParameter.INSTANCE_ELEVATION_PARAM,
            BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,
            BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM,
            BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
            BuiltInParameter.STAIRS_RAILING_HEIGHT_OFFSET };
        
        private WPFDisplayItem GetItemByElement(Element element, string name, string header, string description, string currentlevel, string attachedlevel, Status status, BoundingBoxXYZ box)
        {
            StatusExtended exstatus;
            switch (status)
            {
                case Status.Error:
                    exstatus = StatusExtended.Critical;
                    break;
                default:
                    exstatus = StatusExtended.Warning;
                    break;
            }
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus);
            try
            {
                item.SetZoomParams(element, box);
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", element.Category.Name);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Связанный уровень: ", Description = string.Format("«{0}»", currentlevel) });
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Предложенный уровень: ", Description = string.Format("«{0}»", attachedlevel) });
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Подсказка: ", Description = description });
                HashSet<string> values = new HashSet<string>();
                foreach (BuiltInParameter p in parameters)
                {
                    Parameter parameter = element.get_Parameter(p);
                    if (parameter != null)
                    {
                        if (parameter.StorageType == StorageType.Double)
                        {
                            string value = parameter.AsValueString();

                            if (element.GetType() == typeof(FamilyInstance))
                            {
                                FamilyInstance familyInstance = (FamilyInstance)element;
                                Element host = familyInstance.Host;
                                if (host != null)
                                {
                                    if (host.GetType() == typeof(Floor)) 
                                    {
                                        item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Стоит перенести на корректный уровень", Description = "" });
                                        break;
                                    }
                                }
                            }

                            if (value != null && value != string.Empty && !values.Contains(value))
                            {
                                values.Add(value);
                                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = string.Format("{0}: ", parameter.Definition.Name), Description = string.Format("{0} мм", value) });
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            return item;
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();

                double bpOffset = 0;
                foreach (BasePoint bp in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).ToElements())
                {
                    bpOffset = bp.get_BoundingBox(null).Min.Z;
                }

                string code = null;
                switch (LevelChecker.CheckLevels(commandData.Application.ActiveUIDocument.Document))
                {
                    case CheckResult.NoSections:
                        code = null;
                        break;
                    case CheckResult.Error:
                        TaskDialog TD = new TaskDialog("KPLN: Ошибка");
                        TD.TitleAutoPrefix = false;
                        TD.MainInstruction = "Уровни либо отсутствуют, либо (все либо некоторые из) имеют некорректное наименование.";
                        TD.FooterText = "см. актуальный регламент (обращаться в BIM отдел)";
                        TD.Show();
                        return Result.Failed;
                    case CheckResult.Corpus:
                        code = "К";
                        break;
                    case CheckResult.Sections:
                        code = "С";
                        break;
                }
                
                LevelChecker.LevelCheckers.Clear();
                
                foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
                {
                    LevelChecker.AddLevel(element as Level, doc, code);
                }
                
                List<List<object>> aCats = new List<List<object>>();
                foreach (BuiltInCategory cat in CategoriesToCheck)
                {
                    int ammount = 0;
                    foreach (Element element in new FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType())
                    {
                        if (!element.ElementPassesConditions()) continue;
                        
                        if (element.GetType() == typeof(FamilyInstance))
                        {
                            // Игнорирование вложенных экземпляров в семейство
                            if ((element as FamilyInstance).SuperComponent != null) continue;
                        }
                        
                        try
                        {
                            string Name;
                            #region getName
                            if (element.GetType() == typeof(FamilyInstance))
                            {
                                Name = string.Format("{0}: {1} <{2}>", (element as FamilyInstance).Symbol.FamilyName, (element as FamilyInstance).Symbol.Name, element.Id.ToString());
                            }
                            else
                            {
                                if (element.GetType() == typeof(Wall))
                                {
                                    Name = string.Format("{0}: <{1}>", (element as Wall).WallType.Name, element.Id.ToString());
                                }
                                else
                                {
                                    if (element.GetType() == typeof(Floor))
                                    {
                                        Name = string.Format("{0}: <{1}>", (element as Floor).FloorType.Name, element.Id.ToString());
                                    }
                                    else
                                    {
                                        if (element.GetType() == typeof(Ceiling))
                                        {
                                            Name = string.Format("{0}: <{1}>", (doc.GetElement((element as Ceiling).GetTypeId()) as CeilingType).Name , element.Id.ToString());
                                        }
                                        else
                                        {
                                            Name = string.Format("{0}: <{1}>", element.Name, element.Id.ToString());
                                        }
                                    }
                                }
                            }
                            #endregion
                            
                            CalculateType linkType = CalculateType.Default;
                            Level level = doc.GetElement(element.LevelId) as Level;
                            
                            #region Category and level
                            if (level == null)
                            {
                                try
                                {
                                    level = doc.GetElement(element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).AsElementId()) as Level;
                                }
                                catch (Exception) { }
                            }
                            
                            if (element.GetType() == typeof(FamilyInstance))
                            {
                                FamilyInstance familyInstance = (FamilyInstance)element;
                                if (familyInstance != null)
                                {
                                    Element host = familyInstance.Host;
                                    if (host != null)
                                    {
                                        if (familyInstance.Host.GetType() == typeof(Floor))
                                        {
                                            linkType = CalculateType.Floor;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (element.GetType() == typeof(Floor)) linkType = CalculateType.Floor;
                            }
                            
                            if (level == null) { continue; }
                             
                            #endregion
                            BoundingBoxXYZ box = element.get_BoundingBox(null);
                            // Игнорирование элементов, у которых нет геометрии (например - панели витражей)
                            if (box == null)
                            {
                                continue;
                            }

                            BoundingBoxXYZ boxAnalitical = new BoundingBoxXYZ() { Min = box.Min - new XYZ(0, 0, bpOffset), Max = box.Max - new XYZ(0, 0, bpOffset) };

                            LevelChecker checker = LevelChecker.GetLevelById(level.Id);
                            LevelCheckResult result = linkType == CalculateType.Default ? checker.GetLevelIntersection(boxAnalitical) : checker.GetFloorLevelIntersection(boxAnalitical);
                            LevelChecker c;
                            switch (result)
                            {
                                case LevelCheckResult.NotInside:
                                    if (Check(linkType, LevelCheckResult.NotInside, level, boxAnalitical, doc, out c))
                                    {
                                        outputCollection.Add(GetItemByElement(
                                            element,
                                            Name, 
                                            "[0]: Ошибка",
                                            "Найден более подходящий уровень",
                                            checker.Level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString(),
                                            c.Level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString(),
                                            Status.Error,
                                            box));

                                        ammount++;
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        catch (Exception e) { PrintError(e); }
                    }
                    aCats.Add(new List<object> { ammount, cat });
                }
                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                foreach (List<object> cat in aCats)
                {
                    if ((int)cat[0] == 0) { continue; }
                    Category category = Category.GetCategory(doc, (BuiltInCategory)cat[1]);
                    wpfCategories.Add(new WPFDisplayItem(category.Id.IntegerValue, StatusExtended.Critical) 
                    { 
                        Name = string.Format("{0} ({1})", 
                        category.Name, 
                        ((int)cat[0]).ToString()) 
                    });
                }
                List<WPFDisplayItem> sortedOutputCollection = outputCollection.OrderBy(o => o.Header).ToList();
                ObservableCollection<WPFDisplayItem> wpfElements = new ObservableCollection<WPFDisplayItem>();
                int counter = 1;
                foreach (WPFDisplayItem e in sortedOutputCollection)
                {
                    e.Header = string.Format("{0}# {1}", (counter++).ToString(), e.Header);
                    wpfElements.Add(e);
                }
                if (wpfElements.Count != 0)
                {
                    ElementsOutputExtended form = new ElementsOutputExtended(wpfElements, wpfCategories);
                    form.Show();
                }
                else
                {
                    Print("[Уровни] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }

        public static bool Check(CalculateType linkType, LevelCheckResult result, Level level, BoundingBoxXYZ box, Document doc, out LevelChecker checker)
        {
            foreach (LevelChecker c in LevelChecker.GetOtherLevelById(level.Id))
            {
                LevelCheckResult rslt = linkType == CalculateType.Default ? c.GetLevelIntersection(box) : c.GetFloorLevelIntersection(box);
                
                if (rslt == LevelCheckResult.FullyInside)
                {
                    // Игнорирую для КР привязку элементов к уровню выше (так выдаются спеки)
                    if (doc.Title.ToLower().Contains("_кр_") || doc.Title.ToLower().Contains("_kr_") || doc.Title.ToLower().Contains("_kg_"))
                    {
                        if (level.Id == c.Level.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId())
                        {
                            checker = null;
                            return false;
                        }
                    }
                    checker = c;
                    return true;
                }
            }
            checker = null;
            return false;
        }
    }
}

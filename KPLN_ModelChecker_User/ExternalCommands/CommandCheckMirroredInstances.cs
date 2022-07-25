using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckMirroredInstances : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                
                if (doc.Title.ToUpper().Contains("KR") || doc.Title.ToUpper().Contains("КР"))
                {
                    Print("[Зеркальные] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                    return Result.Succeeded;
                }
                
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                List<List<object>> aCats = new List<List<object>>();
                List<BuiltInCategory> checkCatsList = new List<BuiltInCategory>() { 
                    BuiltInCategory.OST_Doors, 
                    BuiltInCategory.OST_Windows, 
                    BuiltInCategory.OST_CurtainWallPanels,
                    BuiltInCategory.OST_MechanicalEquipment
                };
                List<string> ovvkPassNames = new List<string>()
                {
                    "556_",
                    "557_"
                };

                foreach (BuiltInCategory category in checkCatsList)
                {
                    int ammount = 0;
                    IList<Element> curretCatElems = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(category).WhereElementIsNotElementType().ToElements();

                    // Добавляю фильтрацию по имени семейств. В данном случае - только для оборудования ОВ
                    if (category.Equals(BuiltInCategory.OST_MechanicalEquipment))
                    {
                        curretCatElems = PassesBeginsWithRuleCollection(doc, category, ovvkPassNames);
                        if (curretCatElems.Count() == 0) { continue; }
                    }

                    foreach (Element element in curretCatElems)
                    {
                        try
                        {
                            FamilyInstance instance = element as FamilyInstance;
                            BoundingBoxXYZ box = instance.get_BoundingBox(null);
                            Level instanceLevel = null;
                            try
                            {
                                instanceLevel = doc.GetElement(instance.LevelId) as Level;
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    instanceLevel = doc.GetElement(instance.Host.LevelId) as Level;
                                }
                                catch (Exception)
                                { }
                            }
                            string name = string.Format("{0}: {1} <{2}>", instance.Symbol.FamilyName, instance.Symbol.Name, element.Id.ToString());
                            
                            // Для панелей витража - нужно брать основание - host. И основание дополнительно проверять на поворот - flip
                            // Для панелей витража - нужно брать только окна и двери. Они отсеиваются кодом в имени
                            BuiltInCategory enumCat = (BuiltInCategory)element.Category.Id.IntegerValue;
                            if (enumCat == BuiltInCategory.OST_CurtainWallPanels)
                            {
                                string elName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                                if (elName.StartsWith("135_") && elName.ToLower().Contains("двер") | elName.ToLower().Contains("створк"))
                                {
                                    Wall panelHostWall = instance.Host as Wall;
                                    if (panelHostWall.Flipped)
                                    {
                                        outputCollection.Add(GetItemByElement(element, Status.Error, box));
                                        ammount++;
                                    }
                                }
                            }
                            else
                            {
                                if (instance.Mirrored)
                                {
                                    outputCollection.Add(GetItemByElement(element, Status.Error, box));
                                    ammount++;
                                }
                            }
                        }
                        catch (Exception e) { PrintError(e); }
                    }
                    aCats.Add(new List<object> { ammount, category });
                }
                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                foreach (List<object> cat in aCats)
                {
                    if ((int)cat[0] == 0) { continue; }
                    Category category = Category.GetCategory(doc, (BuiltInCategory)cat[1]);
                    wpfCategories.Add(new WPFDisplayItem(category.Id.IntegerValue, StatusExtended.Critical) { Name = string.Format("{0} ({1})", category.Name, ((int)cat[0]).ToString()) });
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
                    Print("[Зеркальные] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Фильтрация по началу имени
        /// </summary>
        /// <param name="passesNames">Коллекция имен для фильтрации</param>
        /// <param name="coll">Коллекция всех элементов в проекте</param>
        private List<Element> PassesBeginsWithRuleCollection(Document doc, BuiltInCategory cat, List<string> passesNames)
        {
            List<Element> results = new List<Element>();
            foreach (string currentName in passesNames)
            {
                FilteredElementCollector elemColl = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(cat).WhereElementIsNotElementType();
                FilterRule fRule = ParameterFilterRuleFactory.CreateBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                elemColl.WherePasses(eFilter);
                results.AddRange(elemColl.ToElements());
            }
            return results;
        }

        private WPFDisplayItem GetItemByElement(Element element, Status status, BoundingBoxXYZ box)
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
                item.Name = string.Format("{0}: {1}", (element as FamilyInstance).Symbol.FamilyName, (element as FamilyInstance).Symbol.Name);
                item.Header = "Зеркальный элемент";
                if (element.GroupId != ElementId.InvalidElementId)
                {
                    item.Description = "Элемент в группе. Необходимо зайти в группу и вручную исправить замечание";
                }
                else
                {
                    item.Description = "Необходимо исправить замечание вручную";
                }
                item.Category = string.Format("<{0}>", element.Category.Name);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                if (element.LevelId != ElementId.InvalidElementId)
                {
                    string levelname = element.Document.GetElement(element.LevelId).get_Parameter(BuiltInParameter.DATUM_TEXT).AsString();
                    item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Связанный уровень: ", Description = string.Format("«{0}»", levelname) });
                }
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Подсказка: ", Description = item.Description });
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            return item;
        }
    }
}

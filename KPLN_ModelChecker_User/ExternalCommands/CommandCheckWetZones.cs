using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Forms;
using System.Collections.Generic;
using System.Linq;


namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandCheckWetZones : IExternalCommand
    {
        internal const string PluginName = "АР: Проверка мокрых зон";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;


            FilteredElementCollector collectorRooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();

            var windowSPM = new WetZoneParameterWindow(collectorRooms);
            var resultWSPM = windowSPM.ShowDialog();
            string selectedParam = null;

            if (resultWSPM == true)
            {
                selectedParam = windowSPM.SelectedParameter;           
            }
            else
            {
                TaskDialog.Show(PluginName, "Операция отменена.");
                return Result.Cancelled;
            }

            Dictionary<Element, string> roomParamValues = new Dictionary<Element, string>();

            foreach (Element room in collectorRooms)
            {
                Parameter param = room.LookupParameter(selectedParam);
                if (param != null && param.StorageType == StorageType.String)
                {
                    string value = param.AsString() ?? string.Empty;
                    roomParamValues[room] = value;                  
                }
            }

            if (roomParamValues.Count > 0)
            {
                var livingRooms = roomParamValues.Where(r => WetZoneCategories.LivingRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();
                var kitchenRooms = roomParamValues.Where(r => WetZoneCategories.KitchenRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();
                var wetRooms = roomParamValues.Where(r => WetZoneCategories.WetRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();

                var allAssigned = new HashSet<Element>(livingRooms.Concat(kitchenRooms).Concat(wetRooms));
                var undefinedRooms = roomParamValues.Keys.Where(r => !allAssigned.Contains(r)).ToList();

                WetZoneReviewWindow reviewWindow = new WetZoneReviewWindow(uiDoc, doc, selectedParam, livingRooms, kitchenRooms,wetRooms,undefinedRooms);
                reviewWindow.ShowDialog();
            }
            else
            {
                TaskDialog.Show(PluginName, "Подходящие помещения не найдены.");
            }

            return Result.Succeeded;
        }
    }


    ///////////////////////////////////////////
    /// <summary>
    /// Списки названий помещений по категориям
    /// </summary>
    static class WetZoneCategories
    {
        /// Название помещения "Жилые комнаты"
        public static readonly List<string> LivingRooms = new List<string>
        {
            "Жилая комната",
            "Спальня",
            "Гостевая",
            "Гостинная",
            "Гостиная",
            "Детская",
            "Игровая",
            "Кабинет"
        };

        /// Название помещения "Кухни"
        public static readonly List<string> KitchenRooms = new List<string>
        {
            "Кухня",
            "Кухня-ниша",
            "Кухня-столовая",
            "Кухня столовая",
            "Кухонная зона кухни-столовой"
        };

        /// Название помещения "Мокрая зона"
        public static readonly List<string> WetRooms = new List<string>
        {
            "Санузел",
            "С/У",
            "Туалет",
            "Ванная",
            "Ванная комната",
            "Душевая",
            "Постирочная",
            "Прачечная",
        };
    }
}

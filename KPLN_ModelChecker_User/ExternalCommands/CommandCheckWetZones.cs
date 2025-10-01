using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Forms;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
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

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r =>
                {
                    var param = r.LookupParameter("Назначение");
                    return param != null && param.AsString() == "Квартира";
                })
                .ToList();

            var windowSPM = new WetZoneParameterWindow(rooms);
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

            foreach (Element room in rooms)
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
                WetZoneCategories.LoadForDocument(doc);

                var livingRooms = roomParamValues.Where(r => WetZoneCategories.LivingRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();
                var kitchenRooms = roomParamValues.Where(r => WetZoneCategories.KitchenRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();
                var wetRooms = roomParamValues.Where(r => WetZoneCategories.WetRooms.Contains(r.Value.Trim(), System.StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();

                var allAssigned = new HashSet<Element>(livingRooms.Concat(kitchenRooms).Concat(wetRooms));
                var undefinedRooms =  roomParamValues.Where(r => !allAssigned.Contains(r.Key) && !WetZoneCategories.NonProcessedRooms.Contains(r.Value.Trim(), 
                    StringComparer.InvariantCultureIgnoreCase)).Select(r => r.Key).ToList();

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
        public static List<string> LivingRooms { get; private set; } = new List<string>();
        public static List<string> KitchenRooms { get; private set; } = new List<string>();
        public static List<string> WetRooms { get; private set; } = new List<string>();
        public static List<string> NonProcessedRooms { get; private set; } = new List<string>();
        public static List<string> InvalidEquipment { get; private set; } = new List<string>();

        public static readonly string BasePath = @"X:\BIM\6_Инструменты\Плагин мокрые зоны\";
        private const string MainFileName = "_categoriesMain.json";

        public static void LoadForDocument(Document doc)
        {
            string fileName = Path.GetFileNameWithoutExtension(doc.PathName);
            if (string.IsNullOrEmpty(fileName)) throw new Exception("Файл не сохранён, невозможно определить имя модели.");
            string prefix = fileName.Split('_').FirstOrDefault();
            if (string.IsNullOrEmpty(prefix)) throw new Exception("Не удалось извлечь префикс из имени файла.");

            // Загружаем основную базу
            RoomCategoryData baseData = null;
            try
            {
                var mainPath = Path.Combine(BasePath, MainFileName);
                if (!File.Exists(mainPath))
                {
                    TaskDialog.Show("Ошибка", $"Не найден основной JSON-файл определения категорий помещений");
                    return;
                }

                baseData = LoadFromJson(mainPath);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", $"Не удалось загрузить основной JSON-файл определения категорий помещений:\n{ex.Message}");
                return;
            }

            // Загружаем оверрайд, если он есть
            RoomCategoryData overrideData = null;
            try
            {
                var overridePath = Path.Combine(BasePath, $"{prefix}.json");
                if (File.Exists(overridePath))
                    overrideData = LoadFromJson(overridePath);
            }
            catch{}

            ApplyMergedCategories(baseData, overrideData);
        }

        private static void ApplyMergedCategories(RoomCategoryData baseData, RoomCategoryData overrideData)
        {
            Dictionary<string, string> termToCategory = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

            void RegisterTerms(IEnumerable<string> terms, string category)
            {
                if (terms == null) return;

                foreach (var raw in terms)
                {
                    var term = raw?.Trim();
                    if (string.IsNullOrEmpty(term)) continue;

                    if (!termToCategory.ContainsKey(term))
                    {
                        termToCategory[term] = category;
                    }
                }
            }

            RegisterTerms(baseData.LivingRooms, "LivingRooms");
            RegisterTerms(baseData.KitchenRooms, "KitchenRooms");
            RegisterTerms(baseData.WetRooms, "WetRooms");
            RegisterTerms(baseData.NonProcessedRooms, "NonProcessedRooms");
            RegisterTerms(baseData.InvalidEquipment, "InvalidEquipment");

            if (overrideData != null)
            {
                void OverrideTerms(IEnumerable<string> terms, string category)
                {
                    if (terms == null) return;

                    foreach (var raw in terms)
                    {
                        var term = raw?.Trim();
                        if (string.IsNullOrEmpty(term)) continue;

                        if (termToCategory.TryGetValue(term, out string currentCategory))
                        {
                            if (currentCategory == category)
                                continue; 
                        }

                        termToCategory[term] = category;
                    }
                }

                OverrideTerms(overrideData.LivingRooms, "LivingRooms");
                OverrideTerms(overrideData.KitchenRooms, "KitchenRooms");
                OverrideTerms(overrideData.WetRooms, "WetRooms");
                OverrideTerms(overrideData.NonProcessedRooms, "NonProcessedRooms");
                OverrideTerms(overrideData.InvalidEquipment, "InvalidEquipment");
            }

            LivingRooms = new List<string>();
            KitchenRooms = new List<string>();
            WetRooms = new List<string>();
            NonProcessedRooms = new List<string>();
            InvalidEquipment = new List<string>();

            foreach (var kv in termToCategory)
            {
                switch (kv.Value)
                {
                    case "LivingRooms": LivingRooms.Add(kv.Key); break;
                    case "KitchenRooms": KitchenRooms.Add(kv.Key); break;
                    case "WetRooms": WetRooms.Add(kv.Key); break;
                    case "NonProcessedRooms": NonProcessedRooms.Add(kv.Key); break;
                    case "InvalidEquipment": InvalidEquipment.Add(kv.Key); break;
                }
            }
        }

        // загрузка JSON
        private static RoomCategoryData LoadFromJson(string path)
        {
            string json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<RoomCategoryData>(json);
        }

        // Класс только для JSON сериализации
        private class RoomCategoryData
        {
            public List<string> LivingRooms { get; set; }
            public List<string> KitchenRooms { get; set; }
            public List<string> WetRooms { get; set; }
            public List<string> NonProcessedRooms { get; set; }
            public List<string> InvalidEquipment { get; set; }
        }

        // Cбор уникальных InvalidEquipment из всех JSON в папке
        public static List<string> GetAllInvalidEquipment()
        {
            var set = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

            if (!Directory.Exists(BasePath))
                return new List<string>();

            foreach (var path in Directory.EnumerateFiles(BasePath, "*.json"))
            {
                try
                {
                    var data = LoadFromJson(path);
                    if (data?.InvalidEquipment == null) continue;

                    foreach (var raw in data.InvalidEquipment)
                    {
                        var term = raw?.Trim();
                        if (!string.IsNullOrEmpty(term))
                            set.Add(term);
                    }
                }
                catch
                {
                    // глотаем ошибки по отдельным файлам, чтобы не падать из-за одного битого JSON
                }
            }

            return set.OrderBy(s => s).ToList();
        }
    }
}

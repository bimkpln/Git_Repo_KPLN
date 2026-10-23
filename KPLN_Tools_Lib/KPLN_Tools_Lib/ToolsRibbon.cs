using Autodesk.Revit.UI;
using KPLN_Tools_Lib.ExternalCommands;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace KPLN_Tools_Lib
{
    /// <summary>Общие кнопки для выпадающих списков инструментов ЭОМ, ОВВК и СС.</summary>
    public static class ToolsRibbon
    {
        public static PushButton AddTagSpacesButton(PulldownButton pulldown)
            => AddTagButton(pulldown, false);

        public static PushButton AddTagRoomsButton(PulldownButton pulldown)
            => AddTagButton(pulldown, true);

        private static PushButton AddTagButton(PulldownButton pulldown, bool roomsOnly)
        {
            if (pulldown == null)
                throw new ArgumentNullException(nameof(pulldown));

            Assembly assembly = typeof(TagSpacesExtCmd).Assembly;
            var data = new PushButtonData(roomsOnly ? "KPLN_Tools_Lib.TagRooms" : "KPLN_Tools_Lib.TagSpaces",
                roomsOnly ? TagRoomsExtCmd.PluginName : TagSpacesExtCmd.PluginName,
                assembly.Location, roomsOnly ? typeof(TagRoomsExtCmd).FullName : typeof(TagSpacesExtCmd).FullName)
            {
                ToolTip = roomsOnly ? "Маркировать помещения на выбранных планах этажей / потолков." : "Маркировать помещения или пространства на выбранных планах этажей / потолков.",
                LongDescription = (roomsOnly
                    ? "Выделите планы в диспетчере проекта через Ctrl / Shift. Выберите тип марки и источники помещений: текущую модель или загруженные связи.\n"
                    : "Выделите планы в диспетчере проекта через Ctrl / Shift. Выберите категорию, "
                    + "тип марки и источник помещений: текущую модель или загруженную связь. "
                    + "Пространства маркируются в текущей модели. Уже маркированные элементы пропускаются.\n")
                    + string.Format("Дата сборки: {0:yyyy/MM/dd}\nНомер сборки: {1}\nИмя модуля: {2}",
                        File.GetLastWriteTime(assembly.Location), assembly.GetName().Version, assembly.GetName().Name)
            };

            using (Stream stream = assembly.GetManifestResourceStream("KPLN_Tools_Lib.Imagens.tagSpaces.png"))
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.DecodePixelWidth = 16;
                image.DecodePixelHeight = 16;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
                data.Image = image;
                data.LargeImage = image;
            }

            return pulldown.AddPushButton(data);
        }

        public static PushButton AddCreateSpacesButton(PulldownButton pulldown)
        {
            if (pulldown == null)
                throw new ArgumentNullException(nameof(pulldown));

            Assembly assembly = typeof(CreateSpacesExtCmd).Assembly;
            var data = new PushButtonData("KPLN_Tools_Lib.CreateSpaces", CreateSpacesExtCmd.PluginName,
                assembly.Location, typeof(CreateSpacesExtCmd).FullName)
            {
                ToolTip = "Создать пространства по помещениям выбранных связей Revit.",
                LongDescription = "Выбор загруженных связей и стадии пространств. Перенос положения и высоты помещений, "
                    + "сопоставление уровней по отметкам, пропуск существующих пространств.\n"
                    + string.Format("Дата сборки: {0:yyyy/MM/dd}\nНомер сборки: {1}\nИмя модуля: {2}",
                        File.GetLastWriteTime(assembly.Location), assembly.GetName().Version, assembly.GetName().Name)
            };

            using (Stream stream = assembly.GetManifestResourceStream("KPLN_Tools_Lib.Imagens.createSpaces.png"))
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.DecodePixelWidth = 16;
                image.DecodePixelHeight = 16;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
                data.Image = image;
                data.LargeImage = image;
            }

            return pulldown.AddPushButton(data);
        }
    }
}

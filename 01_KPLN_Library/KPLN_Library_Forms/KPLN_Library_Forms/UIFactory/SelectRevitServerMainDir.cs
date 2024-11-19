using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Library_Forms.UIFactory
{
    /// <summary>
    // Инициализация загрузки окна выбора КОРНЕВОЙ папки Revit-Server
    /// </summary>
    public class SelectRevitServerMainDir
    {
        /// <summary>
        /// Метод создания окна для выбора корневой папки сервера
        /// </summary>
        /// <returns>Путь к корневой папки, с учетом имени сервера RS. Формат: "HOSTNAME\PATH"</returns>
        public static ElementSinglePick CreateForm_SelectRSMainDir(int revitVersion)
        {
            ObservableCollection<ElementEntity> rsColl = null;
            switch (revitVersion)
            {
                case 2020:
                    rsColl = new ObservableCollection<ElementEntity>()
                    {
                        new ElementEntity(@"rs02\ФСК_Измайловский_1оч", "Содержит модели АР"),
                        new ElementEntity(@"rs03\Самолет_Сетунь\АР", "Содержит модели АР"),
                        new ElementEntity(@"rs04\Самолет_Сетунь\КР", "Содержит модели КР"),
                        new ElementEntity(@"rs05\Самолет_Сетунь\АУПТ", "Содержит модели АУПТ"),
                        new ElementEntity(@"rs05\Самолет_Сетунь\ВК", "Содержит модели ВК"),
                        new ElementEntity(@"rs05\Самолет_Сетунь\ОВ", "Содержит модели ОВ"),
                        new ElementEntity(@"rs05\Самолет_Сетунь\СС", "Содержит модели СС"),
                        new ElementEntity(@"rs05\Самолет_Сетунь\ЭОМ", "Содержит модели ЭОМ"),
                    };
                    break;
                case 2023:
                    rsColl = new ObservableCollection<ElementEntity>()
                    {
                        new ElementEntity(@"192.168.0.5\Пушкино, Маяковского, 1 очередь", "Содержит модели всех разделов"),
                        new ElementEntity(@"192.168.0.5\Речной порт Якутск", "Содержит модели всех разделов"),
                    };
                    break;
            }

            ElementSinglePick pickForm = new ElementSinglePick(rsColl.OrderBy(p => p.Name), "Выбери корневую папку Revit-Server");

            return pickForm;
        }
    }
}
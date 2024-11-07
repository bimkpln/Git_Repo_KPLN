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
        public static ElementSinglePick CreateForm_SelectRSMainDir()
        {
            ObservableCollection<ElementEntity> rsColl = new ObservableCollection<ElementEntity>()
            {
                new ElementEntity(@"rs02\ФСК_Измайловский_1оч", "Содержит проекты: ИЗМЛ_АР"),
                new ElementEntity(@"rs03\Самолет_Сетунь", "Содержит проекты: СЕТ_1_АР"),
                new ElementEntity(@"rs04\Самолет_Сетунь", "Содержит проекты: СЕТ_1_КР"),
                new ElementEntity(@"rs05\Самолет_Сетунь\АУПТ", "Содержит проекты: СЕТ_1_АУПТ"),
                new ElementEntity(@"rs05\Самолет_Сетунь\ВК", "Содержит проекты: СЕТ_1_ВК"),
                new ElementEntity(@"rs05\Самолет_Сетунь\ОВ", "Содержит проекты: СЕТ_1_ОВ"),
                new ElementEntity(@"rs05\Самолет_Сетунь\СС", "Содержит проекты: СЕТ_1_СС"),
                new ElementEntity(@"rs05\Самолет_Сетунь\ЭОМ", "Содержит проекты: СЕТ_1_ЭОМ"),
            };

            ElementSinglePick pickForm = new ElementSinglePick(rsColl.OrderBy(p => p.Name), "Выбери корневую папку Revit-Server");

            return pickForm;
        }
    }
}

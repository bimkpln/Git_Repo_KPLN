using KPLN_Tools.Common.LinkManager;
using System.Collections.ObjectModel;
using System.Windows;

namespace KPLN_Tools.Forms.Models.Core
{
    public interface IRLinkUserControl
    {
        /// <summary>
        /// Коллекция сущностей для менеджера связей
        /// </summary>
        ObservableCollection<LinkManagerEntity> LinkChangeEntityColl { get; set; }

        /// <summary>
        /// Добавить новую сущность 
        /// </summary>
        void AddNewItem(LinkManagerEntity entity);

        /// <summary>
        /// Удалить текущую сущность (обработка кнопки)
        /// </summary>
        void RemoveItem(object sender, RoutedEventArgs e);
    }
}

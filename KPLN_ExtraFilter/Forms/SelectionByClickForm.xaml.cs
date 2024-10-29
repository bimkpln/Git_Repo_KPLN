using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Entities;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByClickForm : Window
    {
        public SelectionByClickForm(Document doc)
        {
            InitializeComponent();

            CurrentSelectionEntity = new SelectionByClickEntity() { Where_Model = true };

            this.CHB_SameWorkset.IsEnabled = doc.IsWorkshared;
            this.DataContext = CurrentSelectionEntity;
        }

        /// <summary>
        /// Выбранная сущность для запуска
        /// </summary>
        public SelectionByClickEntity CurrentSelectionEntity { get; private set; }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; } = false;

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            Close();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (CurrentSelectionEntity.What_SameCategory
                || CurrentSelectionEntity.What_SameFamily
                || CurrentSelectionEntity.What_SameType
                || CurrentSelectionEntity.What_Workset)
            {
                RunBtn.IsEnabled = true;
            }
            else
            {
                RunBtn.IsEnabled = false;
            }

        }
    }
}

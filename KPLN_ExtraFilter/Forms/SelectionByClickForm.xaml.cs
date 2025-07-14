using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Entities.SelectionByClick;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByClickForm : Window
    {
        private readonly Document _doc;
        private readonly Element _userSelElem;
        
        public SelectionByClickForm(Document doc, Element userSelElem, object lastRunConfigObj)
        {
            _doc = doc;
            _userSelElem = userSelElem;
            InitializeComponent();

            if (lastRunConfigObj != null && lastRunConfigObj is SelectionByClickEntity entity)
            {
                entity.UpdateParams(_doc, _userSelElem);
                
                CurrentSelectionEntity = entity;
                CleareConfigBtn.IsEnabled = true;
            }
            else
                SetDefaultEntity();

            this.CHB_SameWorkset.IsEnabled = doc.IsWorkshared;
            this.DataContext = CurrentSelectionEntity;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Выбранная сущность для запуска
        /// </summary>
        public SelectionByClickEntity CurrentSelectionEntity { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (CurrentSelectionEntity.What_SameCategory
                || CurrentSelectionEntity.What_SameFamily
                || CurrentSelectionEntity.What_SameType
                || CurrentSelectionEntity.What_Workset
                || CurrentSelectionEntity.What_ParameterData)
            {
                RunBtn.IsEnabled = true;
                CleareConfigBtn.IsEnabled = true;
            }
            else
            {
                RunBtn.IsEnabled = false;
                CleareConfigBtn.IsEnabled = false;
            }

        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CleareConfigBtn_Click(object sender, RoutedEventArgs e)
        {
            SetDefaultEntity();
            this.DataContext = CurrentSelectionEntity;
        }

        private void SetDefaultEntity() => 
            CurrentSelectionEntity = new SelectionByClickEntity(_doc, _userSelElem);
    }
}

using KPLN_ExtraFilter.Forms.Entities;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionOptions_MainControl : UserControl
    {
        public static readonly DependencyProperty CurrentSOEntityProperty =
            DependencyProperty.Register(
                nameof(CurrentSOMainControlEntity),
                typeof(SelectionByClickM),
                typeof(SelectionOptions_MainControl),
                new PropertyMetadata(null));

        public SelectionOptions_MainControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Ссылка на общую сущность для управления. 
        /// ВАЖНО: имя в классах VM или на бэке - должно совпадать, чтобы работала DependencyProperty
        /// </summary>
        public SelectionByClickM CurrentSOMainControlEntity
        {
            get => (SelectionByClickM)GetValue(CurrentSOEntityProperty);
            set => SetValue(CurrentSOEntityProperty, value);
        }
    }
}

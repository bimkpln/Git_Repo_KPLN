using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.WPFItems;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System;
using System.Windows;
using System.Windows.Controls;
using KPLN_Library_ExtensibleStorage;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class OutputMainForm : Window
    {
        private ExtensibleStorageBuilder _esBuilderRun;
        private ExtensibleStorageBuilder _esLastText;

        public OutputMainForm(WPFReportCreator creator)
        {
            InitializeComponent();

            this.Title = $"[KPLN]: {creator.CheckName}";
            LastRunData.Text = creator.LogLastRun;

            #region Настраиваю видимость блока ключевого лога
            if (creator.LogMarker is null)
            {
                MarkerRow.Height = new GridLength(0);
                MarkerDataHeader.Visibility = Visibility.Collapsed;
                MarkerData.Visibility = Visibility.Collapsed;
            }
            else
            {
                MarkerRow.Height = GridLength.Auto;
                MarkerData.Text = creator.LogMarker;
            }
            #endregion

            iControll.ItemsSource = creator.WPFEntityCollection;

            cbxFiltration.ItemsSource = creator.WPFFiltration;
            cbxFiltration.SelectedIndex = 0;
        }

        public OutputMainForm(WPFReportCreator creator, ExtensibleStorageBuilder esBuilderRun) : this(creator)
        {
            _esBuilderRun = esBuilderRun;
        }

        public OutputMainForm(WPFReportCreator creator, ExtensibleStorageBuilder esBuilderRun, ExtensibleStorageBuilder esLastText) : this(creator, esBuilderRun)
        {
            _esLastText = esLastText;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ModuleData.CommandQueue.Enqueue(new CommandSetExrStr_TimeRunLog(_esBuilderRun, DateTime.Now));
        }

        private void UpdateCollection(int itemCatId, int itemId)
        {
            foreach (WPFEntity item in iControll.ItemsSource)
            {
                if (itemCatId == -1)
                {
                    item.Visibility = Visibility.Visible;
                }
                else
                {

                }
            }
        }

        private void OnZoomClick(object sender, RoutedEventArgs e)
        {
            WPFEntity wpfEntity = (sender as Button).DataContext as WPFEntity;
            if (wpfEntity != null)
            {
                if (wpfEntity.IsZoomElement)
                    ModuleData.CommandQueue.Enqueue(new CommandZoomElement(wpfEntity.Element, wpfEntity.Box, wpfEntity.Centroid));
                else
                    ModuleData.CommandQueue.Enqueue(new CommandShowElement(wpfEntity.Element));
            }
        }

        private void OnApproveClick(object sender, RoutedEventArgs e)
        {
            WPFEntity wpfEntity = (sender as Button).DataContext as WPFEntity;
            if (wpfEntity != null)
            {
                if (wpfEntity.IsApproveElement)
                {
                    UserTextInput userTextInput = new UserTextInput("Опиши причину");
                    userTextInput.ShowDialog();
                    
                    if (userTextInput.Status == UIStatus.RunStatus.Run)
                    {
                        ModuleData.CommandQueue.Enqueue(new CommandSetExtStr_TextLog(wpfEntity.Element, _esLastText, userTextInput.UserInput));
                    }
                }
            }
        }

        private void OnSelectedCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            //int itemCatId = (cbxFiltration.SelectedItem as WPFEntity).CategoryId;
            //int itemId = (cbxFiltration.SelectedItem as WPFEntity).ElementId;
            //UpdateCollection(itemCatId, itemId);
        }
    }

    ///// <summary>
    ///// Класс для переопределения видимости высоты строк
    ///// </summary>
    //[ValueConversion(typeof(bool), typeof(GridLength))]
    //public class BoolToGridRowHeightConverter_WPFEntityInfo : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        WPFEntity entity = value as WPFEntity;
    //        if (entity != null)
    //        {
    //            if (entity.Info == null)
    //                return ((bool)value == true) ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
    //        }

    //        return value == null ? Visibility.Hidden : Visibility.Visible;
    //    }

    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}

    ///// <summary>
    ///// Класс для переопределения видимости полей WPFEntity в xml
    ///// </summary>
    //public class NullVisibilityConverter_WPFEntityInfo : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        WPFEntity entity = value as WPFEntity;
    //        if (entity != null)
    //        {
    //            if (entity.Info == null)
    //                return Visibility.Hidden;
    //        }

    //        return value == null ? Visibility.Hidden : Visibility.Visible;
    //    }

    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}

    ///// <summary>
    ///// Класс для переопределения видимости высоты строк
    ///// </summary>
    //[ValueConversion(typeof(bool), typeof(GridLength))]
    //public class BoolToGridRowHeightConverter_Marker : IValueConverter
    //{
    //    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        WPFEntity entity = value as WPFEntity;
    //        if (entity != null)
    //        {
    //            if (entity.Info == null)
    //                return ((bool)value == true) ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
    //        }

    //        return value == null ? Visibility.Hidden : Visibility.Visible;
    //    }

    //    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}
}

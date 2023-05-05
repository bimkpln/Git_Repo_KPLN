﻿using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static KPLN_Loader.Output.Output;
using KPLN_ModelChecker_Lib.Common;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для ElementsOutputExtended.xaml
    /// </summary>
    public partial class OutputForm : Window
    {
        private bool _isToggle = true;

        public OutputForm(IList<ElementEntity> collection, IList<ElementEntity> filtration)
        {
            InitializeComponent();
            try
            {
                iControll.ItemsSource = collection;
                cbxFiltration.ItemsSource = filtration;
                cbxFiltration.SelectedIndex = 0;
            }
            catch (Exception e)
            {
                PrintError(e);
            }

        }
        private void UpdateCollection(int itemCatId, int itemId)
        {
            foreach (WPFDisplayItem item in iControll.ItemsSource as ObservableCollection<WPFDisplayItem>)
            {
                if (itemCatId == -1)
                {
                    item.Visibility = Visibility.Visible;
                }
                else if (itemCatId == -2)
                {
                    if (item.ElementId == itemId)
                    {
                        item.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        item.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    if (item.CategoryId == itemCatId)
                    {
                        item.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        item.Visibility = Visibility.Collapsed;
                    }
                }
            }
        }
        private void OnZoomClick(object sender, RoutedEventArgs e)
        {
            WPFDisplayItem element = (sender as Button).DataContext as WPFDisplayItem;
            element.Background = new SolidColorBrush(Color.FromArgb(255, 110, 215, 89));

            // Поиск вида для элементов
            if (element.Box != null)
            {
                ModuleData.CommandQueue.Enqueue(new CommandZoomElement(element.Element, element.Box, element.Centroid));
                TogglerClick(element);
            }
            else
            {
                ModuleData.CommandQueue.Enqueue(new CommandZoomElement(element.Element));
                TogglerClick(element);
            }
        }

        private void OnSelectedCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            int itemCatId = (cbxFiltration.SelectedItem as WPFDisplayItem).CategoryId;
            int itemId = (cbxFiltration.SelectedItem as WPFDisplayItem).ElementId;
            UpdateCollection(itemCatId, itemId);
        }

        private void TogglerClick(WPFDisplayItem element)
        {
            if (_isToggle)
            {
                element.Background = new SolidColorBrush(Color.FromArgb(255, 110, 215, 89));
                _isToggle = false;
            }
            else
            {
                element.Background = new SolidColorBrush(Color.FromArgb(255, 218, 247, 166));
                _isToggle = true;
            }
        }

    }
}

using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class OutputMainForm : Window
    {
        /// <summary>
        /// Extensible Storage: данные по последнему запуску
        /// </summary>
        private readonly ExtensibleStorageBuilder _esBuilderRun;
        /// <summary>
        /// Extensible Storage: данные по пользовательскому комментарию
        /// </summary>
        private readonly ExtensibleStorageBuilder _esBuilderUserText;
        /// <summary>
        /// Extensible Storage: данные по ключевому логу
        /// </summary>
        private readonly ExtensibleStorageBuilder _esBuilderMarker;
        /// <summary>
        /// Revit-application
        /// </summary>
        private readonly UIApplication _application;
        /// <summary>
        /// Ссылка на ExternalCommand, для перезапуска плагина
        /// </summary>
        private readonly string _externalCommand;
        /// <summary>
        /// Коллеция WPFEntity, которая должна отображаться в отчете
        /// </summary>
        private readonly IEnumerable<WPFEntity> _entities;
        private CollectionViewSource _entityViewSource;

        public OutputMainForm(UIApplication uiapp, string externalCommand, WPFReportCreator creator)
        {
            _application = uiapp;
            _externalCommand = externalCommand;
            _entities = creator.WPFEntityCollection;

            InitializeComponent();

            this.Title = $"[KPLN]: {creator.CheckName}";
            LastRunData.Text = creator.LogLastRun;
            cbxFiltration.ItemsSource = creator.FiltrationCollection;

            #region Скрываю видимость блока ключевого лога (он нужен только при использовании спец. конструктора)
            MarkerRow.Height = new GridLength(0);
            MarkerDataHeader.Visibility = Visibility.Collapsed;
            MarkerData.Visibility = Visibility.Collapsed;
            #endregion

            InitializeCollectionViewSource();
            UpdateEntityList();
        }

        public OutputMainForm(UIApplication uiapp, string externalCommand, WPFReportCreator creator, ExtensibleStorageBuilder esBuilderRun, ExtensibleStorageBuilder esBuilderUserText) : this(uiapp, externalCommand, creator)
        {
            _esBuilderRun = esBuilderRun;
            _esBuilderUserText = esBuilderUserText;
        }

        public OutputMainForm(UIApplication uiapp, string externalCommand, WPFReportCreator creator, ExtensibleStorageBuilder esBuilderRun, ExtensibleStorageBuilder esBuilderUserText, ExtensibleStorageBuilder esBuilderMarker) : this(uiapp, externalCommand, creator, esBuilderRun, esBuilderUserText)
        {
            _esBuilderMarker = esBuilderMarker;

            #region Настраиваю данные блока ключевого лога
            MarkerRow.Height = GridLength.Auto;
            MarkerData.Text = creator.LogMarker;
            #endregion
        }

        /// <summary>
        /// Генерация и настройка CollectionViewSource
        /// </summary>
        private void InitializeCollectionViewSource()
        {
            _entityViewSource = new CollectionViewSource { Source = _entities };
            _entityViewSource.Filter += EntityViewSource_Filtered;
            iControll.ItemsSource = _entityViewSource.View;
        }

        /// <summary>
        /// Событие фильтрации для CollectionViewSource
        /// </summary>
        private void EntityViewSource_Filtered(object sender, FilterEventArgs e)
        {
            if (!(e.Item is WPFEntity entity)) return;

            var selectedContent = cbxFiltration.SelectedItem;
            if (selectedContent != null)
            {
                string selectedName = selectedContent.ToString();
                if (chbxApproveShow.IsChecked is true)
                    e.Accepted = selectedName == "Необработанные предупреждения" || entity.FiltrationDescription == selectedName;
                else if (selectedName == "Необработанные предупреждения")
                    e.Accepted = entity.CurrentStatus != Common.Collections.Status.Approve;
                else if (selectedName == "Допустимое")
                    e.Accepted = entity.FiltrationDescription == selectedName;
                else
                    e.Accepted = entity.FiltrationDescription == selectedName && entity.CurrentStatus != Common.Collections.Status.Approve;
            }
        }

        /// <summary>
        /// Обновление данных в окне (CollectionViewSource)
        /// </summary>
        private void UpdateEntityList()
        {
            if (_entityViewSource != null)
                _entityViewSource.View.Refresh();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(_esBuilderRun, DateTime.Now));
        }

        private void OnZoomClicked(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                if (wpfEntity.IsZoomElement)
                    ModuleData.CommandQueue.Enqueue(new CommandZoomElement(wpfEntity.Element, wpfEntity.Box, wpfEntity.Centroid));
                else
                    ModuleData.CommandQueue.Enqueue(new CommandShowElement(wpfEntity.Element));
            }
        }

        private void OnApproveClicked(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                if (wpfEntity.IsApproveElement)
                {
                    UserTextInput userTextInput = new UserTextInput("Опиши причину");
                    userTextInput.ShowDialog();

                    if (userTextInput.Status == UIStatus.RunStatus.Run)
                    {
                        ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetApprComm(wpfEntity, _esBuilderUserText, userTextInput.UserInput));
                    }
                }
            }
        }

        private void OnSelectedCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateEntityList();
        }

        /// <summary>
        /// Перезапустить плагин БЕЗ обновления информации по последнему запуску
        /// </summary>
        private void RestartBtn_Clicked(object sender, RoutedEventArgs e)
        {
            Type type = Type.GetType($"KPLN_ModelChecker_User.ExternalCommands.{_externalCommand}", true);
            AbstrUserOutput instance = Activator.CreateInstance(type) as AbstrUserOutput;
            instance.Execute(_application);

            this.Close();
        }

        private void ChbxApproveShow_Clicked(object sender, RoutedEventArgs e)
        {
            UpdateEntityList();
        }
    }
}

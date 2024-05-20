﻿using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.ExternalCommands;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
        private readonly WPFReportCreator _creator;
        private CollectionViewSource _entityViewSource;

        public OutputMainForm(UIApplication uiapp, string externalCommand, WPFReportCreator creator)
        {
            _application = uiapp;
            _externalCommand = externalCommand;
            _creator = creator;
            _entities = _creator.WPFEntityCollection;

            InitializeComponent();

            this.Title = $"[KPLN]: {creator.CheckName}";
            LastRunData.Text = creator.LogLastRun;
            cbxFiltration.ItemsSource = creator.FiltrationCollection;
            txbCount.Text = _entities.Count().ToString();

            #region Скрываю видимость блока ключевого лога (он нужен только при использовании спец. конструктора)
            MarkerRow.Height = new GridLength(0);
            MarkerDataHeader.Visibility = Visibility.Collapsed;
            MarkerData.Visibility = Visibility.Collapsed;
            #endregion

            InitializeCollectionViewSource();
            UpdateEntityList();

            // Блокирую возможность перезапуска у проверок, которые содержат транзакции (они не открываются вне Ревит) или которые содержат подписки на обработчики событий в конексте Revit API
            if (_externalCommand == nameof(CommandCheckLinks) || _externalCommand == nameof(CommandCheckFamilies)) 
                this.RestartBtn.Visibility = Visibility.Collapsed;
        }

        public OutputMainForm(UIApplication uiapp, string externalCommand, WPFReportCreator creator, ExtensibleStorageBuilder esBuilderRun, ExtensibleStorageBuilder esBuilderUserText, ExtensibleStorageBuilder esBuilderMarker) : this(uiapp, externalCommand, creator)
        {
            _esBuilderRun = esBuilderRun;
            _esBuilderUserText = esBuilderUserText;
            
            #region Настраиваю данные блока ключевого лога
            _esBuilderMarker = esBuilderMarker;
            if (!_esBuilderMarker.Guid.Equals(Guid.Empty))
            {
                MarkerRow.Height = GridLength.Auto;
                MarkerData.Text = creator.LogMarker;
                MarkerDataHeader.Visibility = Visibility.Visible;
                MarkerData.Visibility = Visibility.Visible;
            }
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
                if (selectedName == "Допустимое")
                    e.Accepted = entity.CurrentStatus == Common.CheckCommandCollections.CheckStatus.Approve;
                else if (chbxApproveShow.IsChecked is true)
                    e.Accepted = selectedName == "Необработанные предупреждения" || entity.FiltrationDescription == selectedName;
                else if (selectedName == "Необработанные предупреждения")
                    e.Accepted = entity.CurrentStatus != Common.CheckCommandCollections.CheckStatus.Approve;
                else
                    e.Accepted = entity.FiltrationDescription == selectedName && entity.CurrentStatus != Common.CheckCommandCollections.CheckStatus.Approve;
            }
        }

        /// <summary>
        /// Обновление данных в окне (CollectionViewSource)
        /// </summary>
        private void UpdateEntityList()
        {
            if (_entityViewSource != null)
            {
                _entityViewSource.View.Refresh();
                txbCount.Text = _entityViewSource.View.Cast<WPFEntity>().Count().ToString();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(_esBuilderRun, DateTime.Now));
        }

        private void OnZoomClicked(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                wpfEntity.BackgroundLightening();
                
                if (wpfEntity.IsZoomElement)
                {
                    if (wpfEntity.Element != null)
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandZoomElement(wpfEntity.Element, wpfEntity.Box, wpfEntity.Centroid));
                    else
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandZoomElement(wpfEntity.ElementCollection));
                }
                else
                    if (wpfEntity.Element != null)
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(wpfEntity.Element));
                    else
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(wpfEntity.ElementCollection));
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
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetApprComm(wpfEntity, _esBuilderUserText, userTextInput.UserInput));
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
            // Создаем тип
            Type type = Type.GetType($"KPLN_ModelChecker_User.ExternalCommands.{_externalCommand}", true);

            // Создаем экземпляр типа
            object instance = Activator.CreateInstance(type);
            
            // Определяем метод ExecuteByUIApp
            MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");

            // Вызываем метод ExecuteByUIApp, передавая _uiApp как аргумент
            if (executeMethod != null)
                executeMethod.Invoke(instance, new object[] { _application });
            else
                throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");
                

            this.Close();
        }

        /// <summary>
        /// Экспортировать отчет в Excel
        /// </summary>
        private void ExportBtn_Clicked(object sender, RoutedEventArgs e)
        {
            string path = WPFEntity_ExportToExcel.SetPath();
            if (!string.IsNullOrEmpty(path)) WPFEntity_ExportToExcel.Run(this, path, _creator.CheckName, _entities);
        }

        private void ChbxApproveShow_Clicked(object sender, RoutedEventArgs e)
        {
            UpdateEntityList();
        }
    }
}

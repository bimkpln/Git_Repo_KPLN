using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.ExecutableCommand;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_ModelChecker_Lib;
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
using static System.Net.Mime.MediaTypeNames;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class CheckMainForm : Window
    {
        /// <summary>
        /// Записывать данные по последнему запуску?
        /// </summary>
        private readonly bool _setLastRun;
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
        /// <summary>
        /// Тип плагина, который запускает окно
        /// </summary>
        private readonly Type _pluginType;
        private readonly WPFReportCreator _creator;
        private CollectionViewSource _entityViewSource;

        public CheckMainForm(UIApplication uiapp, string externalCommand, Type pluginType, WPFReportCreator creator, bool setLastRun)
        {
            _application = uiapp;
            _externalCommand = externalCommand;
            _pluginType = pluginType;
            _creator = creator;
            _setLastRun = setLastRun;
            _entities = _creator.WPFEntityCollection;

            InitializeComponent();

            this.Title = $"[KPLN]: {creator.CheckName}";
            LastRunData.Text = creator.LogLastRun;
            cbxFiltration.ItemsSource = creator.FiltrationCollection;
            txbCount.Text = _entities.Count().ToString();

            #region Скрываю видимость блока ключевого лога (он нужен только при использовании спец. конструктора)
            MarkerRow.Height = new GridLength(0);
            MarkerDataHeader.Visibility = System.Windows.Visibility.Collapsed;
            MarkerData.Visibility = System.Windows.Visibility.Collapsed;
            #endregion

            InitializeCollectionViewSource();
            UpdateEntityList();

            // Блокирую возможность перезапуска у проверок, которые содержат транзакции (они не открываются вне Ревит) или которые содержат подписки на обработчики событий в конексте Revit API
            if (_externalCommand == nameof(CommandCheckFamilies)) 
                this.RestartBtn.Visibility = System.Windows.Visibility.Collapsed;
        }

        public CheckMainForm(UIApplication uiapp, string externalCommand, Type pluginType, WPFReportCreator creator, bool setLastRun, ExtensibleStorageBuilder esBuilderRun, ExtensibleStorageBuilder esBuilderUserText, ExtensibleStorageBuilder esBuilderMarker) : this(uiapp, externalCommand, pluginType, creator, setLastRun)
        {
            _esBuilderRun = esBuilderRun;
            _esBuilderUserText = esBuilderUserText;
            
            #region Настраиваю данные блока ключевого лога
            _esBuilderMarker = esBuilderMarker;
            if (!_esBuilderMarker.Guid.Equals(Guid.Empty))
            {
                MarkerRow.Height = GridLength.Auto;
                MarkerData.Text = creator.LogMarker;
                MarkerDataHeader.Visibility = System.Windows.Visibility.Visible;
                MarkerData.Visibility = System.Windows.Visibility.Visible;
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
                    e.Accepted = entity.CurrentStatus == ErrorStatus.Approve;
                else if (chbxApproveShow.IsChecked is true)
                    e.Accepted = selectedName == "Необработанные предупреждения" || entity.FiltrationDescription == selectedName;
                else if (selectedName == "Необработанные предупреждения")
                    e.Accepted = entity.CurrentStatus != ErrorStatus.Approve;
                else
                    e.Accepted = entity.FiltrationDescription == selectedName && entity.CurrentStatus != ErrorStatus.Approve;
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
            if (_setLastRun)
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(_esBuilderRun, DateTime.Now));
        }

        private void OnSelectClicked(object sender, RoutedEventArgs e)
        {
            if((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                if (wpfEntity.Element != null)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSelectElements(new List<Element>(1) { wpfEntity.Element }));
                else
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSelectElements(wpfEntity.ElementCollection));
            }
        }

        private void OnZoomClicked(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                wpfEntity.BackgroundLightening();
                
                if (wpfEntity.IsZoomElement)
                {
                    if (wpfEntity.Element != null)
                    {
                        if (wpfEntity.Box == null || wpfEntity.Centroid == null)
                            HtmlOutput.Print($"Ошибка - у элемента {wpfEntity.Element} не предопределена геометрия. Ищи элемент вручную, через id\n", MessageType.Warning);
                        else
                            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ZoomElementCommand(wpfEntity.Element, wpfEntity.Box, wpfEntity.Centroid));
                    }
                    else
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ZoomElementCommand(wpfEntity.ElementCollection));
                }
                else
                {
                    if (wpfEntity.Element != null)
                    {
                        // Для поиска вида размеров - нельзя использовать обработчики событий, поэтому - через отдельный метод
                        if (wpfEntity.Element is Dimension || wpfEntity.Element is DimensionType) CheckDimension_OpenView.OpenViewForDimensions(_application, wpfEntity.Element);
                        else KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(new List<Element>(1) { wpfEntity.Element }));
                    }
                    else
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(wpfEntity.ElementCollection));
                }
            }
        }

        private void OnApproveClicked(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext is WPFEntity wpfEntity)
            {
                if (wpfEntity.IsApproveElement)
                {
                    UserTextInput userTextInput = new UserTextInput("Опиши причину");
                    if ((bool)userTextInput.ShowDialog())
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetApprComm(wpfEntity, _esBuilderUserText, userTextInput.UserInput));
                }
            }
        }

        private void OnSelectedCategoryChanged(object sender, SelectionChangedEventArgs e) => UpdateEntityList();

        /// <summary>
        /// Перезапустить плагин БЕЗ обновления информации по последнему запуску
        /// </summary>
        private void RestartBtn_Clicked(object sender, RoutedEventArgs e)
        {
            try
            {
                // Создаем тип
                Type type = Type.GetType($"KPLN_ModelChecker_User.ExternalCommands.{_externalCommand}", true);

                // Создаем экземпляр типа
                object instance = Activator.CreateInstance(type);

                // Определяем метод ExecuteByUIApp
                MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");
                MethodInfo executeMethodRef = executeMethod.MakeGenericMethod(_pluginType);

                // Вызываем метод ExecuteByUIApp, передавая _uiApp как аргумент
                if (executeMethod != null)
                    executeMethodRef.Invoke(instance, new object[] { _application, false, false, true, false, true });
                else
                    throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");
            }
            catch (Exception)
            {
                TaskDialog.Show("KPLN: Ошибка", "Не удалось перезапустить. Запусти плагин заново, из меню KPLN");
            }

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

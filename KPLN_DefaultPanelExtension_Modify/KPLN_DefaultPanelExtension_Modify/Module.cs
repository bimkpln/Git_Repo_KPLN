using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_DefaultPanelExtension_Modify.ExecutableCommands;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;

namespace KPLN_DefaultPanelExtension_Modify
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
        private readonly ObservableCollection<Element> _selElems = new ObservableCollection<Element>();
        
        private UIControlledApplication _controlledApp;

        private Autodesk.Windows.RibbonPanel _modifyPanel;
        private Autodesk.Windows.RibbonButton _sendToBtrBtn;
        private const string _sendToBtrBtnId = "ExtCmdSendToBitrix";
        private Autodesk.Windows.RibbonButton _positionBtn;
        private const string _positionBtnId = "ExtCmdListVPPosition";

        public Result Close()
        {
#if !Debug2020 && !Revit2020
            _controlledApp.SelectionChanged -= new EventHandler<SelectionChangedEventArgs>(OnSelectionChanged);
            Autodesk.Windows.ComponentManager.UIElementActivated -= new EventHandler<Autodesk.Windows.UIElementActivatedEventArgs>(OnUiElementActivated);
#endif

            return Result.Succeeded;
        }

        /// <summary>
        /// Добавление кнопки в существующие панели Ревит.Источник: https://jeremytammik.github.io/tbc/a/1170_ts_2_modify_tab_button.htm
        /// ВАЖНО: НЕ РАБОТАЕТ с Ревит2020. Сборка создана для заглушки, чтобы в модуле KPLN_Loader не возникали предупреждения
        /// </summary>
        /// <returns></returns>
        public Result Execute(UIControlledApplication application, string tabName)
        {
#if !Debug2020 && !Revit2020
            _controlledApp = application;
            _controlledApp.SelectionChanged += new EventHandler<SelectionChangedEventArgs>(OnSelectionChanged);
            _selElems.CollectionChanged += SelElems_CollectionChanged;

            // Подписка на событие нажатия кнопки в окне ревит (нужна при добавлении своих кнопок в стандартные панели ревита)
            Autodesk.Windows.ComponentManager.UIElementActivated += new EventHandler<Autodesk.Windows.UIElementActivatedEventArgs>(OnUiElementActivated);

            // Установка основных полей модуля
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Поиск табы Modify
            Autodesk.Windows.RibbonTab currentRTab = null;
            Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
            foreach (Autodesk.Windows.RibbonTab tab in ribbon.Tabs)
            {
                if (tab.Id == "Modify")
                {
                    currentRTab = tab;
                    break;
                }
            }
            if (currentRTab == null) return Result.Cancelled;


            // Создаю свою панель
            _modifyPanel = new Autodesk.Windows.RibbonPanel
            {
                IsVisible = false
            };

            Autodesk.Windows.RibbonPanelSource source = new Autodesk.Windows.RibbonPanelSource
            {
                Name = "KPLN",
                Id = "KPLN",
                Title = "KPLN"
            };

            _modifyPanel.Source = source;
            _modifyPanel.FloatingOrientation = System.Windows.Controls.Orientation.Vertical;


            // Создаю кнопки
            _sendToBtrBtn = CreateButton(
                _sendToBtrBtnId,
                _sendToBtrBtnId,
                $"{ExtCmdSendToBitrix.PluginName}",
                $"{ExtCmdSendToBitrix.PluginName}",
                "Генерируется сообщение с данными по элементу, дополнительными комментариями и отправляется выбранному / -ым пользователям Bitrix.");

            _positionBtn = CreateButton(
                _positionBtnId,
                _positionBtnId,
                $"{ExcCmdListVPPositionStart.PluginName}",
                $"{ExcCmdListVPPositionStart.PluginName}",
                "Сохраняет и применяет положение вида на листе.\nИнструкция допступна на мудл, введите в поиск \"Положение вида\"");



            // В свою панель добавляю кнопки
            _modifyPanel.Source.Items.Add(_sendToBtrBtn);
            _modifyPanel.Source.Items.Add(new Autodesk.Windows.RibbonRowBreak());
            _modifyPanel.Source.Items.Add(_positionBtn);
            _modifyPanel.Source.Items.Add(new Autodesk.Windows.RibbonRowBreak());


            // В системную табу добавляю свою панель
            currentRTab.Panels.Add(_modifyPanel);
#endif
            return Result.Succeeded;
        }

#if !Debug2020 && !Revit2020
        /// <summary>
        /// React to Revit view activation.
        /// </summary>
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is UIApplication uiapp)
            {
                if (uiapp.ActiveUIDocument == null)
                    return;

                _selElems.Clear();
                Document selDoc = uiapp.ActiveUIDocument.Document;
                foreach (ElementId id in e.GetSelectedElements())
                {
                    _selElems.Add(selDoc.GetElement(id));
                }
            }
        }
#endif

        /// <summary>
        /// React to UI element activation, 
        /// e.g. button click. We have absolutely 
        /// no access to the Revit API in this method!
        /// </summary>
        private void OnUiElementActivated(object sender, Autodesk.Windows.UIElementActivatedEventArgs e)
        {
            if(_selElems != null && _selElems.Any())
            {
                if (e.Item?.Id == _sendToBtrBtnId)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExtCmdSendToBitrix());

                if (e.Item?.Id == _positionBtnId)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdListVPPositionStart(_selElems.ToArray()));

            }
        }

        /// <summary>
        /// Create a basic ribbon button with an 
        /// identifying number and an image.
        /// </summary>
        private Autodesk.Windows.RibbonButton CreateButton(string name, string imageName, string text, string ttTitle, string ttContent)
        {
            Autodesk.Windows.RibbonButton button = new Autodesk.Windows.RibbonButton
            {
                Name = name,
                Id = name,
                AllowInStatusBar = true,
                AllowInToolBar = true,
                IsEnabled = true,
                IsToolTipEnabled = true,
                IsVisible = false,
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16),
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32),
                ShowImage = true,
                ShowText = false,
                ShowToolTipOnDisabled = true,
                Text = text,
                Size = Autodesk.Windows.RibbonItemSize.Large,
                ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.NoResize,
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                IsCheckable = true,
                Width = 200,
                // Создаю/устанавливаю описание
                ToolTip = new Autodesk.Windows.RibbonToolTip
                {
                    Title = ttTitle,
                    Content = ttContent,
                    ExpandedContent =
                        $"Дата сборки: {ModuleData.Date}\n" +
                        $"Номер сборки: {ModuleData.Version}\n" +
                        $"Имя модуля: {ModuleData.ModuleName}",
                    IsHelpEnabled = false,
                }
            };

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNWindButtonsForImageReverse.Add((button, _positionBtnId, Assembly.GetExecutingAssembly().GetName().Name));
#endif

            return button;
        }

        /// <summary>
        /// Подписка на изм. коллекции эл-в
        /// </summary>
        private void SelElems_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            bool visibleSentToBtrBtn = null != _selElems && _selElems.Any();
            bool visiblePositionBtn = visibleSentToBtrBtn && _selElems.All(el => el is Viewport);

            _modifyPanel.IsVisible = visibleSentToBtrBtn || visiblePositionBtn;
            _sendToBtrBtn.IsVisible = visibleSentToBtrBtn;
            _positionBtn.IsVisible = visiblePositionBtn;
        }
    }
}

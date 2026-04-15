using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_DefaultPanelExtension_Modify.Commands;
using KPLN_DefaultPanelExtension_Modify.ExecutableCommands;
using KPLN_Loader.Common;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;

namespace KPLN_DefaultPanelExtension_Modify
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
        private readonly ObservableCollection<Element> _selElems = new ObservableCollection<Element>();
        private UIApplication _uiApp;

        private UIControlledApplication _controlledApp;

        private Autodesk.Windows.RibbonPanel _modifyPanel;
        private Autodesk.Windows.RibbonButton _sendToBtrBtn;
        private const string _sendToBtrBtnId = "ExtCmdSendToBitrix";
        private Autodesk.Windows.RibbonButton _positionBtn;
        private const string _positionBtnId = "ExtCmdListVPPosition";
        private Autodesk.Windows.RibbonButton _treeModelBtn;
        private const string _treeModelBtnId = "ExtCmdTreeModel";

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
                $"{CmdSendToBitrix.PluginName}",
                $"{CmdSendToBitrix.PluginName}",
                "Генерируется сообщение с данными по элементу, дополнительными комментариями и отправляется выбранному / -ым пользователям Bitrix.");

            _treeModelBtn = CreateButton(
                _treeModelBtnId,
                _treeModelBtnId,
                $"{ExcCmdTreeModel.PluginName}",
                $"{ExcCmdTreeModel.PluginName}",
                "Создать дерево элементов из выбранных");

            _positionBtn = CreateButton(
                _positionBtnId,
                _positionBtnId,
                $"{ExcCmdListVPPositionStart.PluginName}",
                $"{ExcCmdListVPPositionStart.PluginName}",
                "Сохраняет и применяет положение вида на листе.\nИнструкция допступна на мудл, введите в поиск \"Положение вида\"");



            // В свою панель добавляю кнопки
            _modifyPanel.Source.Items.Add(_sendToBtrBtn);
            _modifyPanel.Source.Items.Add(new Autodesk.Windows.RibbonRowBreak());
            _modifyPanel.Source.Items.Add(_treeModelBtn);
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
                _uiApp = uiapp;
                Document selDoc = _uiApp.ActiveUIDocument.Document;
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
            if (_selElems != null && _selElems.Any())
            {
                if (e.Item?.Id == _sendToBtrBtnId)
                    new CmdSendToBitrix().Execute(_uiApp);

                if (e.Item?.Id == _treeModelBtnId)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdTreeModel());

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
                IsEnabled = true,
                IsToolTipEnabled = true,
                IsVisible = false,
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16),
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32),
                ShowImage = true,
                ShowText = false,
                ShowToolTipOnDisabled = true,
                Text = text,
                IsCheckable = true,
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
            KPLN_Loader.Application.KPLNWindButtonsForImageReverse.Add((button, _sendToBtrBtnId, Assembly.GetExecutingAssembly().GetName().Name));
            KPLN_Loader.Application.KPLNWindButtonsForImageReverse.Add((button, _treeModelBtnId, Assembly.GetExecutingAssembly().GetName().Name));
            KPLN_Loader.Application.KPLNWindButtonsForImageReverse.Add((button, _positionBtnId, Assembly.GetExecutingAssembly().GetName().Name));
#endif

            return button;
        }

        /// <summary>
        /// Подписка на изм. коллекции эл-в
        /// </summary>
        private void SelElems_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            bool visibleSelModelElems = null != _selElems && _selElems.Any();
            bool visibleSelVPorts = visibleSelModelElems && _selElems.All(el => el is Viewport || el is ScheduleSheetInstance);

            // Включаю видимость кнопок
            _modifyPanel.IsVisible = visibleSelModelElems || visibleSelVPorts;
            _sendToBtrBtn.IsVisible = visibleSelModelElems;
            _treeModelBtn.IsVisible = visibleSelModelElems;
            _positionBtn.IsVisible = visibleSelVPorts;


            // Настраиваю размер
            if (visibleSelModelElems && visibleSelVPorts)
            {
                _sendToBtrBtn.Size = Autodesk.Windows.RibbonItemSize.Standard;
                _treeModelBtn.Size = Autodesk.Windows.RibbonItemSize.Standard;
                _positionBtn.Size = Autodesk.Windows.RibbonItemSize.Standard;
            }
            else
            {
                _sendToBtrBtn.Size = Autodesk.Windows.RibbonItemSize.Large;
                _treeModelBtn.Size = Autodesk.Windows.RibbonItemSize.Large;
                _positionBtn.Size = Autodesk.Windows.RibbonItemSize.Large;
            }
        }
    }
}

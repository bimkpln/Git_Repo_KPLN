using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_DBWorker;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.ExternalCommands;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.MainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            ExtCmd_SETLinkChanger.SetStaticEnvironment(application);
            LoadRLI_Service.SetStaticEnvironment(application);
            ExcCmd_LinkChanger_Start.SetStaticEnvironment(application);


            //Ищу или создаю панель
            string panelTName = "Инструменты";
            RibbonPanel panel = null;
            IEnumerable<RibbonPanel> tryTPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelTName);
            if (tryTPanels.Any())
                panel = tryTPanels.FirstOrDefault();
            else
                panel = application.CreateRibbonPanel(tabName, panelTName);

            //Ищу или создаю панель
            string panelMName = "Менеджеры";
            RibbonPanel mPanel = null;
            IEnumerable<RibbonPanel> tryMPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelMName);
            if (tryMPanels.Any())
                mPanel = tryMPanels.FirstOrDefault();
            else
                mPanel = application.CreateRibbonPanel(tabName, panelMName);

            #region Отдельные кнопки (голова)
            PushButtonData autoSaveConfig = CreateBtnData(
            ExtCmd_AutoSaveConfig.PluginName,
            ExtCmd_AutoSaveConfig.PluginName,
            "Настроить автосохранение локальной копии модели. ВАЖНО: синхронизация при этом не происходит, только сохранение локальной копии. " +
            "Если нужно внести изменения в модель из хранилища после вылета Revit - откройте последнюю локальную копию и произведите синхронизацию вручную",
            string.Format(
                "Возможности:\n " +
                    "1. Вкл/выкл функцию автосохранения локальной копии.\n" +
                    "2. Настройка частоты автосохранения.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                ModuleData.Date,
                ModuleData.Version,
                ModuleData.ModuleName
            ),
            typeof(ExtCmd_AutoSaveConfig).FullName,
            "KPLN_Tools.Imagens.diskBig.png",
            "KPLN_Tools.Imagens.diskBig.png",
            "http://moodle/mod/book/view.php?id=502&chapterid=1301#:~:text=%D0%9E%D0%A2%D0%94%D0%95%D0%9B%D0%AC%D0%9D%D0%AB%D0%99%20%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%90%D0%92%D0%A2%D0%9E%D0%A1%D0%9E%D0%A5%D0%A0%D0%90%D0%9D%D0%95%D0%9D%D0%98%D0%95%22",
            true);

            panel.AddItem(autoSaveConfig);

            var ascRI = panel.GetItems().FirstOrDefault(item => item.Name.Equals(ExtCmd_AutoSaveConfig.PluginName));
            SetRIShowText(ascRI, false);
            #endregion


            #region Общие инструменты
            PulldownButton sharedPullDownBtn = CreatePulldownButtonInRibbon("Общие",
                "Общие",
                "Общая коллекция мини-плагинов",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                "toolBox",
                panel,
                false);

            PushButtonData autonumber = CreateBtnData(
                ExtCmd_Autonumber.PluginName,
                ExtCmd_Autonumber.PluginName,
                "Нумерация позици в спецификации на +1 от начального значения",
                string.Format(
                    "Алгоритм запуска:\n" +
                        "1. Выделяем стартовую ячейку спецификации;\n" +
                        "2. Вводим данные, которые указаны в окне.\n\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_Autonumber).FullName,
                "KPLN_Tools.Imagens.autonumberSmall.png",
                "KPLN_Tools.Imagens.autonumberSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=687");

            PushButtonData searchUser = CreateBtnData(
                ExtCmd_SearchRevitUser.PluginName,
                ExtCmd_SearchRevitUser.PluginName,
                "Выдает данные KPLN-пользователя Revit",
                string.Format(
                    "Для поиска введи имя Revit-пользователя.\n" +
                    "\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_SearchRevitUser).FullName,
                "KPLN_Tools.Imagens.searchUserSmall.png",
                "KPLN_Tools.Imagens.searchUserSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1301",
                true);

            PushButtonData tagWiper = CreateBtnData(
                ExtCmd_TagWiper.PluginName,
                ExtCmd_TagWiper.PluginName,
                "УДАЛЯЕТ все марки помещений, которые потеряли основу, а также пытается ОБНОВИТЬ связи маркам помещений",
                string.Format(
                    "Варианты запуска:\n" +
                        "1. Выделить ЛМК листы, чтобы проанализировать размещенные на них виды;\n" +
                        "2. Открыть лист, чтобы проанализировать размещенные на нем виды;\n" +
                        "3. Открыть отдельный вид.\n\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_TagWiper).FullName,
                "KPLN_Tools.Imagens.wipeSmall.png",
                "KPLN_Tools.Imagens.wipeSmall.png",
                "http://moodle");

            PushButtonData monitoringHelper = CreateBtnData(
                ExtCmd_ExtraMonitoring.PluginName,
                ExtCmd_ExtraMonitoring.PluginName,
                "Помощь при копировании и проверке значений парамтеров для элементов с мониторингом",
                string.Format("\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_ExtraMonitoring).FullName,
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "http://moodle");

            PushButtonData changeLevel = CreateBtnData(
                "Изменения позиции элементов на уровне",
                "Изменения позиции элементов на уровне",
                "Плагин для изменения позиции элементов на уровне",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_ChangeLevel).FullName,
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "http://moodle/");

            PushButtonData movingElementsInLevel = CreateBtnData(
                "Перемещение элементов на новый уровень",
                "Перемещение элементов на новый уровень",
                "Плагин для перемещения элементов на новый уровень с сохранением их позиции",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_MovingElementsInLevel).FullName,
                "KPLN_Tools.Imagens.changeElementsInLevelSmall.png",
                "KPLN_Tools.Imagens.changeElementsInLevelSmall.png",
                "http://moodle/");

            // Плагин не реализован до конца. 
            PushButtonData dimensionHelper = CreateBtnData(
                ExtCmd_DimensionHelper.PluginName,
                ExtCmd_DimensionHelper.PluginName,
                "Восстановливает размеры, которые были удалены из-за пересоздания основы",
                string.Format(
                    "Варианты запуска:\n" +
                        "1. Запускаем проект с выгруженной связью и записываем размеры, которые имели к этой связи отношения;\n" +
                        "2. Подгружаем связь, по которой были расставлены размеры. При этом размеры - удаляются (это нормально);\n" +
                        "3. Запускаем плагин и пытаемся восстановить размеры, записанные ранее.\n\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_DimensionHelper).FullName,
                "KPLN_Tools.Imagens.dimHeplerSmall.png",
                "KPLN_Tools.Imagens.dimHeplerSmall.png",
                "http://moodle");

            PushButtonData changeRLinks = CreateBtnData(
                ExtCmd_RLinkManager.PluginName,
                ExtCmd_RLinkManager.PluginName,
                "Загрузить/обновить связи внутри проекта",
                string.Format(
                    "Варианты запуска:\n" +
                        "1. Загрузить связь по указанному пути с сервера KPLN;\n" +
                        "2. Загрузить связь по указанному пути с Revit-Server KPLN;\n" +
                        "3. Обновить связи проекта:\n" +
                        "3.1 Предварительно выделить в диспетчере проекта нужные связи на замену; \n" +
                        "3.2 Просто запустить, тогда все связи появятся в списке на замену. \n\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_RLinkManager).FullName,
                "KPLN_Tools.Imagens.linkChangeSmall.png",
                "KPLN_Tools.Imagens.linkChangeSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1301");

            PushButtonData ws_Links = CreateBtnData(
                    ExtCmd_KR_WSofLinks.PluginName,
                    ExtCmd_KR_WSofLinks.PluginName,
                    "Позволяет включить/выключить рабочий набор, имя которого вы ввели, в связях",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExtCmd_KR_WSofLinks).FullName,
                    "KPLN_Tools.Imagens.wsLinksSmall.png",
                    "KPLN_Tools.Imagens.wsLinksSmall.png",
                    "http://moodle");

#if Revit2020 || Debug2020
            PushButtonData set_ChangeRSLinks = CreateBtnData(
                "СЕТ: Обновить связи",
                "СЕТ: Обновить связи",
                "Обновляет связи между ревит-серверами",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_SETLinkChanger).FullName,
                "KPLN_Tools.Imagens.smlt_Small.png",
                "KPLN_Tools.Imagens.smlt_Small.png",
                "http://moodle");
            sharedPullDownBtn.AddPushButton(set_ChangeRSLinks);
#endif



            sharedPullDownBtn.AddPushButton(autonumber);
            sharedPullDownBtn.AddPushButton(searchUser);
            sharedPullDownBtn.AddPushButton(monitoringHelper);
            sharedPullDownBtn.AddPushButton(tagWiper);
            sharedPullDownBtn.AddPushButton(changeLevel);
            sharedPullDownBtn.AddPushButton(movingElementsInLevel);
            sharedPullDownBtn.AddPushButton(changeRLinks);
            sharedPullDownBtn.AddPushButton(ws_Links);


            #endregion






            #region Отдельные кнопки (хвост)
            // Отверстия только для ИОС
            if (SQLiteMainService.CurrentUserDBSubDepartment.Id != 2 && SQLiteMainService.CurrentUserDBSubDepartment.Id != 3)
            {
                PulldownButton holesPullDownBtn = CreatePulldownButtonInRibbon(
                    "Отверстия",
                    "Отверстия",
                    "Плагины для работы с отверстиями",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    "holes",
                    panel,
                    false);

                PushButtonData holesManagerIOS = CreateBtnData(
                    ExtCmd_HolesManagerIOS.PluginName,
                    ExtCmd_HolesManagerIOS.PluginName,
                    "Подготовка заданий на отверстия от инженеров для АР.",
                    string.Format(
                        "Плагин выполняет следующие функции:\n" +
                            "1. Расширяет специальные элементы семейств, которые позволяют видеть отверстия вне зависимости от секущего диапозона;\n" +
                            "2. Заполняют данные по относительной отметке.\n\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExtCmd_HolesManagerIOS).FullName,
                    "KPLN_Tools.Imagens.holesManagerSmall.png",
                    "KPLN_Tools.Imagens.holesManagerSmall.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1245");

                holesPullDownBtn.AddPushButton(holesManagerIOS);
            }

            // Только для 20 версии, т.к.для более новых появилось событие изменения выбора пользователем
#if Revit2020 || Debug2020
            PushButtonData sendMsgToBitrix = CreateBtnData(
                ExtCmd_SendMsgToBitrix.PluginName,
                ExtCmd_SendMsgToBitrix.PluginName,
                "Отправляет данные по выделенному элементу пользователю в Bitrix",
                string.Format(
                    "Генерируется сообщение с данными по элементу, дополнительными комментариями и отправляется выбранному/-ым пользователям Bitrix.\n" +
                    "\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_SendMsgToBitrix).FullName,
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "http://moodle");
            sendMsgToBitrix.AvailabilityClassName = typeof(ButtonAvailable_UserSelect).FullName;

            panel.AddItem(sendMsgToBitrix);
#endif


            PushButtonData nodeManager = CreateBtnData(
                    ExtCmd_NodeManager.PluginName,
                    ExtCmd_NodeManager.PluginName,
                    "Каталог узлов KPLN",
                    string.Format(
                        "Каталог узлов KPLN.\n" +
                        "\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExtCmd_NodeManager).FullName,
                    "KPLN_Tools.Imagens.nodeManagerBig.png",
                    "KPLN_Tools.Imagens.nodeManagerBig.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1342");

            mPanel.AddItem(nodeManager);
            #endregion

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для создания PushButtonData будущей кнопки
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private PushButtonData CreateBtnData(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            string smlImageName,
            string lrgImageName,
            string contextualHelp,
            bool avclass = false)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className)
            {
                Text = text,
                ToolTip = shortDescription,
                LongDescription = longDescription,
            };
            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            data.Image = PngImageSource(smlImageName);
            data.LargeImage = PngImageSource(lrgImageName);


            if (avclass)
                data.AvailabilityClassName = typeof(StaticAvailable).FullName;

            return data;
        }

        /// <summary>
        /// Метод для добавления иконки ButtonData
        /// </summary>
        /// <param name="embeddedPathname">Имя иконки. Для иконок указать Build Action -> Embedded Resource</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        /// <summary>
        /// Метод для создания PulldownButton из RibbonItem (выпадающий список).
        /// Данный метод добавляет 1 отдельный элемент. Для добавления нескольких - нужны перегрузки методов AddStackedItems (добавит 2-3 элемента в столбик)
        /// </summary>
        /// <param name="name">Внутреннее имя вып. списка</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="imageName">Имя картинки</param>
        private PulldownButton CreatePulldownButtonInRibbon(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string imageName,
            RibbonPanel panel,
            bool showName)
        {
            PulldownButton pullDownRI = panel.AddItem(new PulldownButtonData(name, text)
            {
                ToolTip = shortDescription,
                LongDescription = longDescription,
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16),
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32),
            }) as PulldownButton;

            SetRIShowText(pullDownRI, showName);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((pullDownRI, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif

            return pullDownRI;
        }

        /// <summary>
        /// Тонкая настройка видимости текста RibbonItem
        /// </summary>
        private static void SetRIShowText(RibbonItem ri, bool showName)
        {
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(ri.GetId());
            revitRibbonItem.ShowText = showName;
        }
    }
}
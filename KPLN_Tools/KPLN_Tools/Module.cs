using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.ExternalCommands;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            Command_SETLinkChanger.SetStaticEnvironment(application);
            LoadRLI_Service.SetStaticEnvironment(application);
            CommandLinkChanger_Start.SetStaticEnvironment(application);

            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Инструменты");

            //Добавляю выпадающий список pullDown
            #region Общие инструменты
            PulldownButton sharedPullDownBtn = CreatePulldownButtonInRibbon("Общие",
                "Общие",
                "Общая коллекция мини-плагинов",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                PngImageSource("KPLN_Tools.Imagens.toolBoxSmall.png"),
                PngImageSource("KPLN_Tools.Imagens.toolBoxBig.png"),
                panel,
                false);

            PushButtonData autonumber = CreateBtnData(
                CommandAutonumber.PluginName,
                CommandAutonumber.PluginName,
                "Нумерация позици в спецификации на +1 от начального значения",
                string.Format(
                    "Алгоритм запуска:\n" +
                        "1. Запускаем плагин для фиксации размеров штампов;\n" +
                        "2. Меняем семейство на согласованное с BIM-отделом;\n" +
                        "3. Запускаем плагин для установки размеров листов и добавления параметров.\n\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandAutonumber).FullName,
                "KPLN_Tools.Imagens.autonumberSmall.png",
                "KPLN_Tools.Imagens.autonumberSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=687");

            PushButtonData searchUser = CreateBtnData(
                CommandSearchRevitUser.PluginName,
                CommandSearchRevitUser.PluginName,
                "Выдает данные KPLN-пользователя Revit",
                string.Format(
                    "Для поиска введи имя Revit-пользователя.\n" +
                    "\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandSearchRevitUser).FullName,
                "KPLN_Tools.Imagens.searchUserSmall.png",
                "KPLN_Tools.Imagens.searchUserSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1301",
                true);

            PushButtonData tagWiper = CreateBtnData(
                CommandTagWiper.PluginName,
                CommandTagWiper.PluginName,
                "УДАЛЯЕТ все марки помещений, которые потеряли основу, а также пытается ОБНОВИТЬ связи маркам помещений",
                string.Format(
                    "Варианты запуска:\n" +
                        "1. Выделить ЛМК листы, чтобы проанализировать рамзещенные на них виды;\n" +
                        "2. Открыть лист, чтобы проанализировать рамзещенные на нем виды;\n" +
                        "3. Открыть отдельный вид.\n\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandTagWiper).FullName,
                "KPLN_Tools.Imagens.wipeSmall.png",
                "KPLN_Tools.Imagens.wipeSmall.png",
                "http://moodle");

            PushButtonData monitoringHelper = CreateBtnData(
                CommandExtraMonitoring.PluginName,
                CommandExtraMonitoring.PluginName,
                "Помощь при копировании и проверке значений парамтеров для элементов с мониторингом",
                string.Format("\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandExtraMonitoring).FullName,
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "http://moodle");

            PushButtonData changeLevel = CreateBtnData(
                "Изменение уровня",
                "Изменение уровня",
                "Плагин для изменения позиции уровня с сохранением привязанности элементов",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandChangeLevel).FullName,
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "http://moodle/");

            // Плагин не реализован до конца. 
            PushButtonData dimensionHelper = CreateBtnData(
                CommandDimensionHelper.PluginName,
                CommandDimensionHelper.PluginName,
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
                typeof(CommandDimensionHelper).FullName,
                "KPLN_Tools.Imagens.dimHeplerSmall.png",
                "KPLN_Tools.Imagens.dimHeplerSmall.png",
                "http://moodle");

            PushButtonData changeRLinks = CreateBtnData(
                CommandRLinkManager.PluginName,
                CommandRLinkManager.PluginName,
                "Загрузить/обновлят связи внутри проекта",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandRLinkManager).FullName,
                "KPLN_Tools.Imagens.linkChangeSmall.png",
                "KPLN_Tools.Imagens.linkChangeSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1301");

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
                typeof(Command_SETLinkChanger).FullName,
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
            sharedPullDownBtn.AddPushButton(changeRLinks);
            #endregion

            #region Инструменты АР
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 2 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton arToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Плагины АР",
                    "Плагины АР",
                    "АР: Коллекция плагинов для автоматизации задач",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.arMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.arMainBig.png"),
                    panel,
                    false);

                PushButtonData arGNSArea = CreateBtnData(
                    "Площадь ГНС",
                    "Площадь ГНС",
                    "Обводит внешние границы здания на плане",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_AR_GNSBound).FullName,
                    "KPLN_Tools.Imagens.gnsAreaBig.png",
                    "KPLN_Tools.Imagens.gnsAreaSmall.png",
                    "http://moodle");

                PushButtonData TEPDesign = CreateBtnData(
                    "Оформление ТЭП",
                    "Оформление ТЭП",
                    "Плагин для оформления ТЭП",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_AR_TEPDesign).FullName,
                    "KPLN_Tools.Imagens.TEPDesignBig.png",
                    "KPLN_Tools.Imagens.TEPDesignSmall.png",
                    "http://moodle");

                arToolsPullDownBtn.AddPushButton(arGNSArea);
                arToolsPullDownBtn.AddPushButton(TEPDesign);
            }

            #endregion

            #region Инструменты КР
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 3 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton krToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Плагины КР",
                    "Плагины КР",
                    "КР: Коллекция плагинов для автоматизации задач",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.krMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.krMainBig.png"),
                    panel,
                    false);

                PushButtonData smnx_Rebar = CreateBtnData(
                    "SMNX_Металоёмкость",
                    "SMNX_Металоёмкость",
                    "SMNX: Заполняет параметр \"SMNX_Расход арматуры (Кг/м3)\"",
                    string.Format(
                        "Варианты запуска:\n" +
                            "1. Записать объём бетона и основную марку в арматуру;\n" +
                            "2. Перенести значения из спецификации в параметр;\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_KR_SMNX_RebarHelper).FullName,
                    "KPLN_Tools.Imagens.wipeSmall.png",
                    "KPLN_Tools.Imagens.wipeSmall.png",
                    "http://moodle");

                krToolsPullDownBtn.AddPushButton(smnx_Rebar);
            }
            #endregion

            #region Инструменты ОВВК
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 4
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 5
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ovvkToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Плагины ОВВК",
                    "Плагины ОВВК",
                    "ОВВК: Коллекция плагинов для автоматизации задач",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.hvacSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.hvacBig.png"),
                    panel,
                false);

                PushButtonData ovvk_pipeThickness = CreateBtnData(
                    Command_OVVK_PipeThickness.PluginName,
                    Command_OVVK_PipeThickness.PluginName,
                    "Заполняет толщину труб по выбранной конфигурации",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_OVVK_PipeThickness).FullName,
                    "KPLN_Tools.Imagens.pipeThicknessSmall.png",
                    "KPLN_Tools.Imagens.pipeThicknessSmall.png",
                    "http://moodle");

                PushButtonData ovvk_systemManager = CreateBtnData(
                    Command_OVVK_SystemManager.PluginName,
                    Command_OVVK_SystemManager.PluginName,
                    "Управление системами в проекте",
                    string.Format(
                        "Функционал:" +
                            "\n1. Обновляет имя систем;" +
                            "\n2. Объединяет системы в группы для специфицирования и генерации видов;" +
                            "\n3. Генерация видов." +
                            "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_OVVK_SystemManager).FullName,
                    "KPLN_Tools.Imagens.systemMangerSmall.png",
                    "KPLN_Tools.Imagens.systemMangerSmall.png",
                    "http://moodle");

                PushButtonData ov_ductThickness = CreateBtnData(
                    Command_OV_DuctThickness.PluginName,
                    Command_OV_DuctThickness.PluginName,
                    "Заполняет толщину воздуховодов в зависимости от типа системы и наличия изоляцияя/огнезащиты согласно СП.60 и СП.7",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_OV_DuctThickness).FullName,
                    "KPLN_Tools.Imagens.ductThicknessSmall.png",
                    "KPLN_Tools.Imagens.ductThicknessSmall.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1301");

                PushButtonData ov_ozkDuctAccessory = CreateBtnData(
                    Command_OV_OZKDuctAccessory.PluginName,
                    Command_OV_OZKDuctAccessory.PluginName,
                    "Заполняет данные по ОЗК клапанам",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_OV_OZKDuctAccessory).FullName,
                    "KPLN_Tools.Imagens.ozkDuctAccessorySmall.png",
                    "KPLN_Tools.Imagens.ozkDuctAccessorySmall.png",
                    "http://moodle");

#if Revit2020 || Debug2020
                PushButtonData set_InsulationPipes = CreateBtnData(
                    "ОВВК: СЕТ_Изоляция",
                    "ОВВК: СЕТ_Изоляция",
                    "(ИСПРАВЛЕННАЯ ВЕРСИЯ СМЛТ): Заполняет данные по изоляции",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_SET_InsulationPipes).FullName,
                    "KPLN_Tools.Imagens.smlt_Small.png",
                    "KPLN_Tools.Imagens.smlt_Small.png",
                    "http://moodle");

                ovvkToolsPullDownBtn.AddPushButton(set_InsulationPipes);
#endif

                ovvkToolsPullDownBtn.AddPushButton(ovvk_pipeThickness);
                ovvkToolsPullDownBtn.AddPushButton(ov_ductThickness);
                ovvkToolsPullDownBtn.AddPushButton(ov_ozkDuctAccessory);
                ovvkToolsPullDownBtn.AddPushButton(ovvk_systemManager);
            }
            #endregion

            #region Инструменты СС
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 7 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ssToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Плагины СС",
                    "Плагины СС",
                    "СС: Коллекция плагинов для автоматизации задач",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.ssMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.ssMainBig.png"),
                    panel,
                    false);

                PushButtonData ssSystems = CreateBtnData(
                    "Слаботочные системы",
                    "Слаботочные системы",
                    "Помощь в создании цепей СС",
                    string.Format("Плагин создаёт цепи нестандартным путём - генерируются отдельные системы на участки между 2мя элементами. " +
                        "При этом элементу №2 в качестве щита присваивается элемент №1.\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_SS_Systems).FullName,
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "http://moodle");

                PushButtonData ssFillInParameters = CreateBtnData(
                    "Заполнить параметры на чертежном виде",
                    "Заполнить параметры на чертежном виде",
                    "Заполнить параметры на чертежном виде",
                    string.Format("Плагин заполняет параметр ``КП_Позиция_Сумма`` для одинаковых семейств на чертежном виде, собирая значения параметров ``КП_О_Позиция`` с учетом параметра ``КП_О_Группирование``, " +
                    "а также заполняет параметр ``КП_И_Количество в спецификацию`` для семейств категории ``Элементы узлов`` на чертежном виде, у которых в спецификации необходимо учитывать длину, а не количество\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_FillInParametersSS).FullName,
                    "KPLN_Tools.Imagens.FillInParamSmall.png",
                    "KPLN_Tools.Imagens.FillInParamSmall.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1319");

                ssToolsPullDownBtn.AddPushButton(ssSystems);
                ssToolsPullDownBtn.AddPushButton(ssFillInParameters);
            }
            #endregion

            #region Инструменты ЭОМ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 6 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton eomToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Плагины ЭОМ",
                    "Плагины ЭОМ",
                    "ЭОМ: Коллекция плагинов для автоматизации задач",
                    string.Format(
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.eomMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.eomMainBig.png"),
                    panel,
                    false);

                PushButtonData setParams = CreateBtnData(
                    "СЕТ: Заполнить параметры",
                    "СЕТ: Заполнить параметры",
                    "СЕТ: Заполнить параметры",
                    string.Format("Плагин заполняет параметр для формирования спецификации для: \n" +
                        "1. Кабельных лотков;\n" +
                        "2. Соед. деталей кабельных лотков;\n" +
                        "3. Воздуховодов (огнезащита).\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_SET_EOMParams).FullName,
                    "KPLN_Tools.Imagens.FillInParamSmall.png",
                    "KPLN_Tools.Imagens.FillInParamSmall.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1319");

                eomToolsPullDownBtn.AddPushButton(setParams);
            }

            #endregion


            #region Отверстия
            // Наполняю плагинами в зависимости от отдела
            if (DBWorkerService.CurrentDBUserSubDepartment.Id != 2 && DBWorkerService.CurrentDBUserSubDepartment.Id != 3)
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
                    PngImageSource("KPLN_Tools.Imagens.holesSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.holesBig.png"),
                    panel,
                    false);

                PushButtonData holesManagerIOS = CreateBtnData(
                    CommandHolesManagerIOS.PluginName,
                    CommandHolesManagerIOS.PluginName,
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
                    typeof(CommandHolesManagerIOS).FullName,
                    "KPLN_Tools.Imagens.holesManagerSmall.png",
                    "KPLN_Tools.Imagens.holesManagerSmall.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1245");

                holesPullDownBtn.AddPushButton(holesManagerIOS);
            }
            #endregion

            #region Отдельные кнопки
            PushButtonData sendMsgToBitrix = CreateBtnData(
                CommandSendMsgToBitrix.PluginName,
                CommandSendMsgToBitrix.PluginName,
                "Отправляет данные по выделенному элементу пользователю в Bitrix",
                string.Format(
                    "Генерируется сообщение с данными по элементу, дополнительными комментариями и отправляется выбранному/-ым пользователям Bitrix.\n" +
                    "\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandSendMsgToBitrix).FullName,
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "http://moodle");
            sendMsgToBitrix.AvailabilityClassName = typeof(ButtonAvailable_UserSelect).FullName;

            panel.AddItem(sendMsgToBitrix);
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
                LongDescription = longDescription
            };
            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            data.Image = PngImageSource(smlImageName);
            data.LargeImage = PngImageSource(lrgImageName);
            if (avclass)
            {
                data.AvailabilityClassName = typeof(StaticAvailable).FullName;
            }

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
        /// <param name="imgSmall">Картинка маленькая</param>
        /// <param name="imgBig">Картинка большая</param>
        private PulldownButton CreatePulldownButtonInRibbon(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            ImageSource imgSmall,
            ImageSource imgBig,
            RibbonPanel panel,
            bool showName)
        {
            RibbonItem pullDownRI = panel.AddItem(new PulldownButtonData(name, text)
            {
                ToolTip = shortDescription,
                LongDescription = longDescription,
                Image = imgSmall,
                LargeImage = imgBig,
            });

            // Тонкая настройка видимости RibbonItem
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(pullDownRI.GetId());
            revitRibbonItem.ShowText = showName;

            return pullDownRI as PulldownButton;
        }
    }
}




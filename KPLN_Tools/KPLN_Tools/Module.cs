using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.Docking;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.ExternalCommands;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
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
            ModuleData.MainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);
            
            Command_SETLinkChanger.SetStaticEnvironment(application);
            LoadRLI_Service.SetStaticEnvironment(application);
            CommandLinkChanger_Start.SetStaticEnvironment(application);

            //Äîáàâëÿþ ïàíåëü
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Èíñòðóìåíòû");

            //Äîáàâëÿþ âûïàäàþùèé ñïèñîê pullDown
            #region Îáùèå èíñòðóìåíòû
            PulldownButton sharedPullDownBtn = CreatePulldownButtonInRibbon("Îáùèå",
                "Îáùèå",
                "Îáùàÿ êîëëåêöèÿ ìèíè-ïëàãèíîâ",
                string.Format(
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "Íóìåðàöèÿ ïîçèöè â ñïåöèôèêàöèè íà +1 îò íà÷àëüíîãî çíà÷åíèÿ",
                string.Format(
                    "Àëãîðèòì çàïóñêà:\n" +
                        "1. Çàïóñêàåì ïëàãèí äëÿ ôèêñàöèè ðàçìåðîâ øòàìïîâ;\n" +
                        "2. Ìåíÿåì ñåìåéñòâî íà ñîãëàñîâàííîå ñ BIM-îòäåëîì;\n" +
                        "3. Çàïóñêàåì ïëàãèí äëÿ óñòàíîâêè ðàçìåðîâ ëèñòîâ è äîáàâëåíèÿ ïàðàìåòðîâ.\n\n" +
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "Âûäàåò äàííûå KPLN-ïîëüçîâàòåëÿ Revit",
                string.Format(
                    "Äëÿ ïîèñêà ââåäè èìÿ Revit-ïîëüçîâàòåëÿ.\n" +
                    "\n" +
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "ÓÄÀËßÅÒ âñå ìàðêè ïîìåùåíèé, êîòîðûå ïîòåðÿëè îñíîâó, à òàêæå ïûòàåòñÿ ÎÁÍÎÂÈÒÜ ñâÿçè ìàðêàì ïîìåùåíèé",
                string.Format(
                    "Âàðèàíòû çàïóñêà:\n" +
                        "1. Âûäåëèòü ËÌÊ ëèñòû, ÷òîáû ïðîàíàëèçèðîâàòü ðàçìåùåííûå íà íèõ âèäû;\n" +
                        "2. Îòêðûòü ëèñò, ÷òîáû ïðîàíàëèçèðîâàòü ðàçìåùåííûå íà íåì âèäû;\n" +
                        "3. Îòêðûòü îòäåëüíûé âèä.\n\n" +
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "Ïîìîùü ïðè êîïèðîâàíèè è ïðîâåðêå çíà÷åíèé ïàðàìòåðîâ äëÿ ýëåìåíòîâ ñ ìîíèòîðèíãîì",
                string.Format("\nÄàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandExtraMonitoring).FullName,
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "http://moodle");

            PushButtonData changeLevel = CreateBtnData(
                "Èçìåíåíèå óðîâíÿ",
                "Èçìåíåíèå óðîâíÿ",
                "Ïëàãèí äëÿ èçìåíåíèÿ ïîçèöèè óðîâíÿ ñ ñîõðàíåíèåì ïðèâÿçàííîñòè ýëåìåíòîâ",
                string.Format(
                    "\nÄàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandChangeLevel).FullName,
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "http://moodle/");

            // Ïëàãèí íå ðåàëèçîâàí äî êîíöà. 
            PushButtonData dimensionHelper = CreateBtnData(
                CommandDimensionHelper.PluginName,
                CommandDimensionHelper.PluginName,
                "Âîññòàíîâëèâàåò ðàçìåðû, êîòîðûå áûëè óäàëåíû èç-çà ïåðåñîçäàíèÿ îñíîâû",
                string.Format(
                    "Âàðèàíòû çàïóñêà:\n" +
                        "1. Çàïóñêàåì ïðîåêò ñ âûãðóæåííîé ñâÿçüþ è çàïèñûâàåì ðàçìåðû, êîòîðûå èìåëè ê ýòîé ñâÿçè îòíîøåíèÿ;\n" +
                        "2. Ïîäãðóæàåì ñâÿçü, ïî êîòîðîé áûëè ðàññòàâëåíû ðàçìåðû. Ïðè ýòîì ðàçìåðû - óäàëÿþòñÿ (ýòî íîðìàëüíî);\n" +
                        "3. Çàïóñêàåì ïëàãèí è ïûòàåìñÿ âîññòàíîâèòü ðàçìåðû, çàïèñàííûå ðàíåå.\n\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "Çàãðóçèòü/îáíîâèòü ñâÿçè âíóòðè ïðîåêòà",
                string.Format(
                    "Âàðèàíòû çàïóñêà:\n" +
                        "1. Çàãðóçèòü ñâÿçü ïî óêàçàííîìó ïóòè ñ ñåðâåðà KPLN;\n" +
                        "2. Çàãðóçèòü ñâÿçü ïî óêàçàííîìó ïóòè ñ Revit-Server KPLN;\n" +
                        "3. Îáíîâèòü ñâÿçè ïðîåêòà:\n" +
                        "3.1 Ïðåäâàðèòåëüíî âûäåëèòü â äèñïåò÷åðå ïðîåêòà íóæíûå ñâÿçè íà çàìåíó; \n" +
                        "3.2 Ïðîñòî çàïóñòèòü, òîãäà âñå ñâÿçè ïîÿâÿòñÿ â ñïèñêå íà çàìåíó. \n\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                "ÑÅÒ: Îáíîâèòü ñâÿçè",
                "ÑÅÒ: Îáíîâèòü ñâÿçè",
                "Îáíîâëÿåò ñâÿçè ìåæäó ðåâèò-ñåðâåðàìè",
                string.Format(
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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

            #region Èíñòðóìåíòû ÀÐ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 2 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton arToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Ïëàãèíû ÀÐ",
                    "Ïëàãèíû ÀÐ",
                    "ÀÐ: Êîëëåêöèÿ ïëàãèíîâ äëÿ àâòîìàòèçàöèè çàäà÷",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.arMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.arMainBig.png"),
                    panel,
                    false);

                PushButtonData arGNSArea = CreateBtnData(
                    "Ïëîùàäü ÃÍÑ",
                    "Ïëîùàäü ÃÍÑ",
                    "Îáâîäèò âíåøíèå ãðàíèöû çäàíèÿ íà ïëàíå",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_AR_GNSBound).FullName,
                    "KPLN_Tools.Imagens.gnsAreaBig.png",
                    "KPLN_Tools.Imagens.gnsAreaSmall.png",
                    "http://moodle");

                PushButtonData TEPDesign = CreateBtnData(
                    "Îôîðìëåíèå ÒÝÏ",
                    "Îôîðìëåíèå ÒÝÏ",
                    "Ïëàãèí äëÿ îôîðìëåíèÿ ÒÝÏ",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_AR_TEPDesign).FullName,
                    "KPLN_Tools.Imagens.TEPDesignBig.png",
                    "KPLN_Tools.Imagens.TEPDesignSmall.png",
                    "http://moodle");

                arToolsPullDownBtn.AddPushButton(arGNSArea);
#if Debug2023 || Revit2023
                arToolsPullDownBtn.AddPushButton(TEPDesign);
#endif
            }

            #endregion

            #region Èíñòðóìåíòû ÊÐ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 3 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton krToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Ïëàãèíû ÊÐ",
                    "Ïëàãèíû ÊÐ",
                    "ÊÐ: Êîëëåêöèÿ ïëàãèíîâ äëÿ àâòîìàòèçàöèè çàäà÷",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.krMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.krMainBig.png"),
                    panel,
                    false);

                PushButtonData smnx_Rebar = CreateBtnData(
                    "SMNX_Ìåòàëî¸ìêîñòü",
                    "SMNX_Ìåòàëî¸ìêîñòü",
                    "SMNX: Çàïîëíÿåò ïàðàìåòð \"SMNX_Ðàñõîä àðìàòóðû (Êã/ì3)\"",
                    string.Format(
                        "Âàðèàíòû çàïóñêà:\n" +
                            "1. Çàïèñàòü îáú¸ì áåòîíà è îñíîâíóþ ìàðêó â àðìàòóðó;\n" +
                            "2. Ïåðåíåñòè çíà÷åíèÿ èç ñïåöèôèêàöèè â ïàðàìåòð;\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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

            #region Èíñòðóìåíòû ÎÂÂÊ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 4
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 5
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ovvkToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Ïëàãèíû ÎÂÂÊ",
                    "Ïëàãèíû ÎÂÂÊ",
                    "ÎÂÂÊ: Êîëëåêöèÿ ïëàãèíîâ äëÿ àâòîìàòèçàöèè çàäà÷",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "Çàïîëíÿåò òîëùèíó òðóá ïî âûáðàííîé êîíôèãóðàöèè",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "Óïðàâëåíèå ñèñòåìàìè â ïðîåêòå",
                    string.Format(
                        "Ôóíêöèîíàë:" +
                            "\n1. Îáíîâëÿåò èìÿ ñèñòåì;" +
                            "\n2. Îáúåäèíÿåò ñèñòåìû â ãðóïïû äëÿ ñïåöèôèöèðîâàíèÿ è ãåíåðàöèè âèäîâ;" +
                            "\n3. Ãåíåðàöèÿ âèäîâ." +
                            "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "Çàïîëíÿåò òîëùèíó âîçäóõîâîäîâ â çàâèñèìîñòè îò òèïà ñèñòåìû è íàëè÷èÿ èçîëÿöèÿÿ/îãíåçàùèòû ñîãëàñíî ÑÏ.60 è ÑÏ.7",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "Çàïîëíÿåò äàííûå ïî ÎÇÊ êëàïàíàì",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "ÎÂÂÊ: ÑÅÒ_Èçîëÿöèÿ",
                    "ÎÂÂÊ: ÑÅÒ_Èçîëÿöèÿ",
                    "(ÈÑÏÐÀÂËÅÍÍÀß ÂÅÐÑÈß ÑÌËÒ): Çàïîëíÿåò äàííûå ïî èçîëÿöèè",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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

            #region Èíñòðóìåíòû ÑÑ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 7 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ssToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Ïëàãèíû ÑÑ",
                    "Ïëàãèíû ÑÑ",
                    "ÑÑ: Êîëëåêöèÿ ïëàãèíîâ äëÿ àâòîìàòèçàöèè çàäà÷",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.ssMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.ssMainBig.png"),
                    panel,
                    false);

                PushButtonData ssSystems = CreateBtnData(
                    "Ñëàáîòî÷íûå ñèñòåìû",
                    "Ñëàáîòî÷íûå ñèñòåìû",
                    "Ïîìîùü â ñîçäàíèè öåïåé ÑÑ",
                    string.Format("Ïëàãèí ñîçäà¸ò öåïè íåñòàíäàðòíûì ïóò¸ì - ãåíåðèðóþòñÿ îòäåëüíûå ñèñòåìû íà ó÷àñòêè ìåæäó 2ìÿ ýëåìåíòàìè. " +
                        "Ïðè ýòîì ýëåìåíòó ¹2 â êà÷åñòâå ùèòà ïðèñâàèâàåòñÿ ýëåìåíò ¹1.\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_SS_Systems).FullName,
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "http://moodle");

                PushButtonData ssFillInParameters = CreateBtnData(
                    "Çàïîëíèòü ïàðàìåòðû íà ÷åðòåæíîì âèäå",
                    "Çàïîëíèòü ïàðàìåòðû íà ÷åðòåæíîì âèäå",
                    "Çàïîëíèòü ïàðàìåòðû íà ÷åðòåæíîì âèäå",
                    string.Format("Ïëàãèí çàïîëíÿåò ïàðàìåòð ``ÊÏ_Ïîçèöèÿ_Ñóììà`` äëÿ îäèíàêîâûõ ñåìåéñòâ íà ÷åðòåæíîì âèäå, ñîáèðàÿ çíà÷åíèÿ ïàðàìåòðîâ ``ÊÏ_Î_Ïîçèöèÿ`` ñ ó÷åòîì ïàðàìåòðà ``ÊÏ_Î_Ãðóïïèðîâàíèå``, " +
                    "à òàêæå çàïîëíÿåò ïàðàìåòð ``ÊÏ_È_Êîëè÷åñòâî â ñïåöèôèêàöèþ`` äëÿ ñåìåéñòâ êàòåãîðèè ``Ýëåìåíòû óçëîâ`` íà ÷åðòåæíîì âèäå, ó êîòîðûõ â ñïåöèôèêàöèè íåîáõîäèìî ó÷èòûâàòü äëèíó, à íå êîëè÷åñòâî\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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

            #region Èíñòðóìåíòû ÝÎÌ
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 6 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton eomToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "Ïëàãèíû ÝÎÌ",
                    "Ïëàãèíû ÝÎÌ",
                    "ÝÎÌ: Êîëëåêöèÿ ïëàãèíîâ äëÿ àâòîìàòèçàöèè çàäà÷",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.eomMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.eomMainBig.png"),
                    panel,
                    false);

                PushButtonData setParams = CreateBtnData(
                    "ÑÅÒ: Çàïîëíèòü ïàðàìåòðû",
                    "ÑÅÒ: Çàïîëíèòü ïàðàìåòðû",
                    "ÑÅÒ: Çàïîëíèòü ïàðàìåòðû",
                    string.Format("Ïëàãèí çàïîëíÿåò ïàðàìåòð äëÿ ôîðìèðîâàíèÿ ñïåöèôèêàöèè äëÿ: \n" +
                        "1. Êàáåëüíûõ ëîòêîâ;\n" +
                        "2. Ñîåä. äåòàëåé êàáåëüíûõ ëîòêîâ;\n" +
                        "3. Âîçäóõîâîäîâ (îãíåçàùèòà).\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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


            #region Îòâåðñòèÿ
            // Íàïîëíÿþ ïëàãèíàìè â çàâèñèìîñòè îò îòäåëà
            if (DBWorkerService.CurrentDBUserSubDepartment.Id != 2 && DBWorkerService.CurrentDBUserSubDepartment.Id != 3)
            {
                PulldownButton holesPullDownBtn = CreatePulldownButtonInRibbon(
                    "Îòâåðñòèÿ",
                    "Îòâåðñòèÿ",
                    "Ïëàãèíû äëÿ ðàáîòû ñ îòâåðñòèÿìè",
                    string.Format(
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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
                    "Ïîäãîòîâêà çàäàíèé íà îòâåðñòèÿ îò èíæåíåðîâ äëÿ ÀÐ.",
                    string.Format(
                        "Ïëàãèí âûïîëíÿåò ñëåäóþùèå ôóíêöèè:\n" +
                            "1. Ðàñøèðÿåò ñïåöèàëüíûå ýëåìåíòû ñåìåéñòâ, êîòîðûå ïîçâîëÿþò âèäåòü îòâåðñòèÿ âíå çàâèñèìîñòè îò ñåêóùåãî äèàïîçîíà;\n" +
                            "2. Çàïîëíÿþò äàííûå ïî îòíîñèòåëüíîé îòìåòêå.\n\n" +
                        "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
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

            #region Îòäåëüíûå êíîïêè
            PushButtonData sendMsgToBitrix = CreateBtnData(
                CommandSendMsgToBitrix.PluginName,
                CommandSendMsgToBitrix.PluginName,
                "Îòïðàâëÿåò äàííûå ïî âûäåëåííîìó ýëåìåíòó ïîëüçîâàòåëþ â Bitrix",
                string.Format(
                    "Ãåíåðèðóåòñÿ ñîîáùåíèå ñ äàííûìè ïî ýëåìåíòó, äîïîëíèòåëüíûìè êîììåíòàðèÿìè è îòïðàâëÿåòñÿ âûáðàííîìó/-ûì ïîëüçîâàòåëÿì Bitrix.\n" +
                    "\n" +
                    "Äàòà ñáîðêè: {0}\nÍîìåð ñáîðêè: {1}\nÈìÿ ìîäóëÿ: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandSendMsgToBitrix).FullName,
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "KPLN_Tools.Imagens.sendMsgBig.png",
                "http://moodle");
            sendMsgToBitrix.AvailabilityClassName = typeof(ButtonAvailable_UserSelect).FullName;
            
            PushButtonData familyManagerPanel = CreateBtnData(
                CommandFamilyManager.PluginName,
                CommandFamilyManager.PluginName,
                "ÐÐ°ÑÐ°Ð»Ð¾Ð³ ÑÐµÐ¼ÐµÐ¹ÑÑÐ² KPLN",
                string.Format(
                    "ÐÐ°ÑÐ°Ð»Ð¾Ð³ ÑÐµÐ¼ÐµÐ¹ÑÑÐ² KPLN.\n" +
                    "\n" +
                    "ÐÐ°ÑÐ° ÑÐ±Ð¾ÑÐºÐ¸: {0}\nÐÐ¾Ð¼ÐµÑ ÑÐ±Ð¾ÑÐºÐ¸: {1}\nÐÐ¼Ñ Ð¼Ð¾Ð´ÑÐ»Ñ: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandFamilyManager).FullName,
                "KPLN_Tools.Imagens.familyManagerBig.png",
                "KPLN_Tools.Imagens.familyManagerBig.png",
                "http://moodle");
         
            panel.AddItem(sendMsgToBitrix);
            panel.AddItem(familyManagerPanel);
            #endregion


            return Result.Succeeded;
        }

        /// <summary>
        /// Ìåòîä äëÿ ñîçäàíèÿ PushButtonData áóäóùåé êíîïêè
        /// </summary>
        /// <param name="name">Âíóòðåííåå èìÿ êíîïêè</param>
        /// <param name="text">Èìÿ, âèäèìîå ïîëüçîâàòåëþ</param>
        /// <param name="shortDescription">Êðàòêîå îïèñàíèå, âèäèìîå ïîëüçîâàòåëþ</param>
        /// <param name="longDescription">Ïîëíîå îïèñàíèå, âèäèìîå ïîëüçîâàòåëþ ïðè çàëåðæêå êóðñîðà</param>
        /// <param name="className">Èìÿ êëàññà, ñîäåðæàùåãî ðåàëèçàöèþ êîìàíäû</param>
        /// <param name="contextualHelp">Ññûëêà íà web-ñòðàíèöó ïî êëàâèøå F1</param>
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
        /// Ìåòîä äëÿ äîáàâëåíèÿ èêîíêè ButtonData
        /// </summary>
        /// <param name="embeddedPathname">Èìÿ èêîíêè. Äëÿ èêîíîê óêàçàòü Build Action -> Embedded Resource</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        /// <summary>
        /// Ìåòîä äëÿ ñîçäàíèÿ PulldownButton èç RibbonItem (âûïàäàþùèé ñïèñîê).
        /// Äàííûé ìåòîä äîáàâëÿåò 1 îòäåëüíûé ýëåìåíò. Äëÿ äîáàâëåíèÿ íåñêîëüêèõ - íóæíû ïåðåãðóçêè ìåòîäîâ AddStackedItems (äîáàâèò 2-3 ýëåìåíòà â ñòîëáèê)
        /// </summary>
        /// <param name="name">Âíóòðåííåå èìÿ âûï. ñïèñêà</param>
        /// <param name="text">Èìÿ, âèäèìîå ïîëüçîâàòåëþ</param>
        /// <param name="shortDescription">Êðàòêîå îïèñàíèå, âèäèìîå ïîëüçîâàòåëþ</param>
        /// <param name="longDescription">Ïîëíîå îïèñàíèå, âèäèìîå ïîëüçîâàòåëþ ïðè çàëåðæêå êóðñîðà</param>
        /// <param name="imgSmall">Êàðòèíêà ìàëåíüêàÿ</param>
        /// <param name="imgBig">Êàðòèíêà áîëüøàÿ</param>
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

            // Òîíêàÿ íàñòðîéêà âèäèìîñòè RibbonItem
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(pullDownRI.GetId());
            revitRibbonItem.ShowText = showName;

            return pullDownRI as PulldownButton;
        }
    }
}



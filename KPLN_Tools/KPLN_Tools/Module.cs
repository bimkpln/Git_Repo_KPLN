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
        public static string CurrentRevitVersion;
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            CurrentRevitVersion = application.ControlledApplication.VersionNumber;
            
            Command_SETLinkChanger.SetStaticEnvironment(application);
            LoadRLI_Service.SetStaticEnvironment(application);
            CommandLinkChanger_Start.SetStaticEnvironment(application);

            //�������� ������
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "�����������");

            //�������� ���������� ������ pullDown
            #region ����� �����������
            PulldownButton sharedPullDownBtn = CreatePulldownButtonInRibbon("�����",
                "�����",
                "����� ��������� ����-��������",
                string.Format(
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "��������� ������ � ������������ �� +1 �� ���������� ��������",
                string.Format(
                    "�������� �������:\n" +
                        "1. ��������� ������ ��� �������� �������� �������;\n" +
                        "2. ������ ��������� �� ������������� � BIM-�������;\n" +
                        "3. ��������� ������ ��� ��������� �������� ������ � ���������� ����������.\n\n" +
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "������ ������ KPLN-������������ Revit",
                string.Format(
                    "��� ������ ����� ��� Revit-������������.\n" +
                    "\n" +
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "������� ��� ����� ���������, ������� �������� ������, � ����� �������� �������� ����� ������ ���������",
                string.Format(
                    "�������� �������:\n" +
                        "1. �������� ��� �����, ����� ���������������� ����������� �� ��� ����;\n" +
                        "2. ������� ����, ����� ���������������� ����������� �� ��� ����;\n" +
                        "3. ������� ��������� ���.\n\n" +
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "������ ��� ����������� � �������� �������� ���������� ��� ��������� � ������������",
                string.Format("\n���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandExtraMonitoring).FullName,
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "KPLN_Tools.Imagens.monitorMainSmall.png",
                "http://moodle");

            PushButtonData changeLevel = CreateBtnData(
                "��������� ������",
                "��������� ������",
                "������ ��� ��������� ������� ������ � ����������� ������������� ���������",
                string.Format(
                    "\n���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandChangeLevel).FullName,
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "KPLN_Tools.Imagens.changeLevelSmall.png",
                "http://moodle/");

            // ������ �� ���������� �� �����. 
            PushButtonData dimensionHelper = CreateBtnData(
                CommandDimensionHelper.PluginName,
                CommandDimensionHelper.PluginName,
                "��������������� �������, ������� ���� ������� ��-�� ������������ ������",
                string.Format(
                    "�������� �������:\n" +
                        "1. ��������� ������ � ����������� ������ � ���������� �������, ������� ����� � ���� ����� ���������;\n" +
                        "2. ���������� �����, �� ������� ���� ����������� �������. ��� ���� ������� - ��������� (��� ���������);\n" +
                        "3. ��������� ������ � �������� ������������ �������, ���������� �����.\n\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "���������/�������� ����� ������ �������",
                string.Format(
                    "�������� �������:\n" +
                        "1. ��������� ����� �� ���������� ���� � ������� KPLN;\n" +
                        "2. ��������� ����� �� ���������� ���� � Revit-Server KPLN;\n" +
                        "3. �������� ����� �������:\n" +
                        "3.1 �������������� �������� � ���������� ������� ������ ����� �� ������; \n" +
                        "3.2 ������ ���������, ����� ��� ����� �������� � ������ �� ������. \n\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                "���: �������� �����",
                "���: �������� �����",
                "��������� ����� ����� �����-���������",
                string.Format(
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ����������� ��
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 2 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton arToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "������� ��",
                    "������� ��",
                    "��: ��������� �������� ��� ������������� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.arMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.arMainBig.png"),
                    panel,
                    false);

                PushButtonData arGNSArea = CreateBtnData(
                    "������� ���",
                    "������� ���",
                    "������� ������� ������� ������ �� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_AR_GNSBound).FullName,
                    "KPLN_Tools.Imagens.gnsAreaBig.png",
                    "KPLN_Tools.Imagens.gnsAreaSmall.png",
                    "http://moodle");

                PushButtonData TEPDesign = CreateBtnData(
                    "���������� ���",
                    "���������� ���",
                    "������ ��� ���������� ���",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ����������� ��
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 3 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton krToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "������� ��",
                    "������� ��",
                    "��: ��������� �������� ��� ������������� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.krMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.krMainBig.png"),
                    panel,
                    false);

                PushButtonData smnx_Rebar = CreateBtnData(
                    "SMNX_������������",
                    "SMNX_������������",
                    "SMNX: ��������� �������� \"SMNX_������ �������� (��/�3)\"",
                    string.Format(
                        "�������� �������:\n" +
                            "1. �������� ����� ������ � �������� ����� � ��������;\n" +
                            "2. ��������� �������� �� ������������ � ��������;\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ����������� ����
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 4
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 5
                || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ovvkToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "������� ����",
                    "������� ����",
                    "����: ��������� �������� ��� ������������� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "��������� ������� ���� �� ��������� ������������",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "���������� ��������� � �������",
                    string.Format(
                        "����������:" +
                            "\n1. ��������� ��� ������;" +
                            "\n2. ���������� ������� � ������ ��� ���������������� � ��������� �����;" +
                            "\n3. ��������� �����." +
                            "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "��������� ������� ������������ � ����������� �� ���� ������� � ������� ���������/���������� �������� ��.60 � ��.7",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "��������� ������ �� ��� ��������",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "����: ���_��������",
                    "����: ���_��������",
                    "(������������ ������ ����): ��������� ������ �� ��������",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ����������� ��
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 7 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton ssToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "������� ��",
                    "������� ��",
                    "��: ��������� �������� ��� ������������� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.ssMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.ssMainBig.png"),
                    panel,
                    false);

                PushButtonData ssSystems = CreateBtnData(
                    "����������� �������",
                    "����������� �������",
                    "������ � �������� ����� ��",
                    string.Format("������ ������ ���� ������������� ���� - ������������ ��������� ������� �� ������� ����� 2�� ����������. " +
                        "��� ���� �������� �2 � �������� ���� ������������� ������� �1.\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(Command_SS_Systems).FullName,
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "KPLN_Tools.Imagens.ssSystemsSmall.png",
                    "http://moodle");

                PushButtonData ssFillInParameters = CreateBtnData(
                    "��������� ��������� �� ��������� ����",
                    "��������� ��������� �� ��������� ����",
                    "��������� ��������� �� ��������� ����",
                    string.Format("������ ��������� �������� ``��_�������_�����`` ��� ���������� �������� �� ��������� ����, ������� �������� ���������� ``��_�_�������`` � ������ ��������� ``��_�_�������������``, " +
                    "� ����� ��������� �������� ``��_�_���������� � ������������`` ��� �������� ��������� ``�������� �����`` �� ��������� ����, � ������� � ������������ ���������� ��������� �����, � �� ����������\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ����������� ���
            if (DBWorkerService.CurrentDBUserSubDepartment.Id == 6 || DBWorkerService.CurrentDBUserSubDepartment.Id == 8)
            {
                PulldownButton eomToolsPullDownBtn = CreatePulldownButtonInRibbon(
                    "������� ���",
                    "������� ���",
                    "���: ��������� �������� ��� ������������� �����",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    PngImageSource("KPLN_Tools.Imagens.eomMainSmall.png"),
                    PngImageSource("KPLN_Tools.Imagens.eomMainBig.png"),
                    panel,
                    false);

                PushButtonData setParams = CreateBtnData(
                    "���: ��������� ���������",
                    "���: ��������� ���������",
                    "���: ��������� ���������",
                    string.Format("������ ��������� �������� ��� ������������ ������������ ���: \n" +
                        "1. ��������� ������;\n" +
                        "2. ����. ������� ��������� ������;\n" +
                        "3. ������������ (����������).\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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


            #region ���������
            // �������� ��������� � ����������� �� ������
            if (DBWorkerService.CurrentDBUserSubDepartment.Id != 2 && DBWorkerService.CurrentDBUserSubDepartment.Id != 3)
            {
                PulldownButton holesPullDownBtn = CreatePulldownButtonInRibbon(
                    "���������",
                    "���������",
                    "������� ��� ������ � �����������",
                    string.Format(
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
                    "���������� ������� �� ��������� �� ��������� ��� ��.",
                    string.Format(
                        "������ ��������� ��������� �������:\n" +
                            "1. ��������� ����������� �������� ��������, ������� ��������� ������ ��������� ��� ����������� �� �������� ���������;\n" +
                            "2. ��������� ������ �� ������������� �������.\n\n" +
                        "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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

            #region ��������� ������
            PushButtonData sendMsgToBitrix = CreateBtnData(
                CommandSendMsgToBitrix.PluginName,
                CommandSendMsgToBitrix.PluginName,
                "���������� ������ �� ����������� �������� ������������ � Bitrix",
                string.Format(
                    "������������ ��������� � ������� �� ��������, ��������������� ������������� � ������������ ����������/-�� ������������� Bitrix.\n" +
                    "\n" +
                    "���� ������: {0}\n����� ������: {1}\n��� ������: {2}",
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
        /// ����� ��� �������� PushButtonData ������� ������
        /// </summary>
        /// <param name="name">���������� ��� ������</param>
        /// <param name="text">���, ������� ������������</param>
        /// <param name="shortDescription">������� ��������, ������� ������������</param>
        /// <param name="longDescription">������ ��������, ������� ������������ ��� �������� �������</param>
        /// <param name="className">��� ������, ����������� ���������� �������</param>
        /// <param name="contextualHelp">������ �� web-�������� �� ������� F1</param>
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
        /// ����� ��� ���������� ������ ButtonData
        /// </summary>
        /// <param name="embeddedPathname">��� ������. ��� ������ ������� Build Action -> Embedded Resource</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        /// <summary>
        /// ����� ��� �������� PulldownButton �� RibbonItem (���������� ������).
        /// ������ ����� ��������� 1 ��������� �������. ��� ���������� ���������� - ����� ���������� ������� AddStackedItems (������� 2-3 �������� � �������)
        /// </summary>
        /// <param name="name">���������� ��� ���. ������</param>
        /// <param name="text">���, ������� ������������</param>
        /// <param name="shortDescription">������� ��������, ������� ������������</param>
        /// <param name="longDescription">������ ��������, ������� ������������ ��� �������� �������</param>
        /// <param name="imgSmall">�������� ���������</param>
        /// <param name="imgBig">�������� �������</param>
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

            // ������ ��������� ��������� RibbonItem
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(pullDownRI.GetId());
            revitRibbonItem.ShowText = showName;

            return pullDownRI as PulldownButton;
        }
    }
}



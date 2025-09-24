using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Linq;
using System.Windows.Interop;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    public abstract class AbstrCommand
    {
        /// <summary>
        /// Конструктор для классов, наследуемых от AbstrCheckCommand. Если его не переопределить в наследнике - IExternalCommand не справиться с запуском (ему нужен конструтор по умолчанию)
        /// </summary>
        public AbstrCommand(){ }

        /// <summary>
        /// Коллекция элементов для проверки
        /// </summary>
        public Element[] ElemsToCheck { get; private protected set; }

        /// <summary>
        /// Коллекция элементов с ошибками
        /// </summary>
        public CheckerEntity[] CheckerEntities { get; private protected set; }

        /// <summary>
        /// Ссылка на проверку
        /// </summary>
        public AbstrCheck CommandCheck { get; set; }

        /// <summary>
        /// Спец. метод для вызова данного класса из кнопки WPF: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        public virtual void ExecuteByUIApp<T>(UIApplication uiapp, bool onlyErrorType = false, bool setPluginActivity = false, bool showMainForm = false, bool setLastRun = false, bool showSuccsessText = false) 
            where T : AbstrCheck, new()
        {
            try
            {
                Document doc = uiapp.ActiveUIDocument.Document;

                if (CommandCheck == null || ElemsToCheck == null)
                {
                    CommandCheck = new T().Set_UIAppData(uiapp, doc);
                    ElemsToCheck = CommandCheck.GetElemsToCheck();
                }

                if (setPluginActivity)
                    DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{CommandCheck.PluginName}", ModuleData.ModuleName).ConfigureAwait(false);


                CheckerEntities = CommandCheck.ExecuteCheck(ElemsToCheck, onlyErrorType);
                if (CheckerEntities != null && CheckerEntities.Length > 0 && showMainForm)
                    ReportCreatorAndDemonstrator<T>(uiapp, setLastRun);
                else if (showSuccsessText)
                {
                    // Логируем последний запуск (отдельно, если все было ОК, а потом всплыли ошибки)
                    if (showMainForm)
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(CommandCheck.ESEntity.ESBuilderRun, DateTime.Now));

                    // Выводим, что всё ок
                    HtmlOutput.Print($"[{CommandCheck.ESEntity.CheckName}] Предупреждений не найдено :)", MessageType.Success);
                }
            }
            catch (Exception ex) 
            {
                HtmlOutput.Print($"[{CommandCheck.ESEntity.CheckName}] - ошибка при проверке: {ex.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Подготовка окна результата проверки для пользователя
        /// </summary>
        /// <param name="uiapp">Revit-UIApplication</param>
        /// <param name="setLastRun">Нужно ли записывать последний запуск?</param>
        /// <param name="isMarkered">Нужно ли использовать основной маркер при создании окна?</param>
        /// <returns>Окно для вывода пользователю</returns>
        public void ReportCreatorAndDemonstrator<T>(UIApplication uiapp, bool setLastRun = false, bool isMarkered = false) where T : AbstrCheck, new()
        {
            WPFReportCreator repCreator = CreateReport(uiapp.ActiveUIDocument.Document, isMarkered);
            SetWPFEntityFiltration(repCreator);

            CheckMainForm form = new CheckMainForm(uiapp, this.GetType().Name, typeof(T), repCreator, setLastRun, CommandCheck.ESEntity.ESBuilderRun, CommandCheck.ESEntity.ESBuilderUserText, CommandCheck.ESEntity.ESBuildergMarker);
            
            // Связываю с окном ревит, откуда был запуск
            IntPtr windHandle = CommandCheck.CheckUIApp.MainWindowHandle;
            new WindowInteropHelper(form)
            {
                Owner = windHandle
            };

            form.Show();
        }

        /// <summary>
        /// Установить фильтрацию элементов в отчете
        /// </summary>
        private protected virtual void SetWPFEntityFiltration(WPFReportCreator report) => report.SetWPFEntityFiltration_ByErrorHeader();

        /// <summary>
        /// Метод для подготовки отчета
        /// </summary>
        private WPFReportCreator CreateReport(Document doc, bool isMarkered)
        {
            #region Настройка информации по логам проека
            Element piElem = doc.ProjectInformation;
            ResultMessage esMsgRun = CommandCheck.ESEntity.ESBuilderRun.GetResMessage_Element(piElem);
            ResultMessage esMsgMarker = CommandCheck.ESEntity.ESBuildergMarker.GetResMessage_Element(piElem);
            #endregion

            WPFReportCreator result = null;
            WPFEntity[] wpfEntityColl = CheckerEntities.Select(chEnt => new WPFEntity(chEnt, CommandCheck.ESEntity)).ToArray();
            if (CommandCheck.ESEntity.ESBuildergMarker.Guid.Equals(Guid.Empty))
                result = new WPFReportCreator(wpfEntityColl, CommandCheck.ESEntity.CheckName, esMsgRun.Description);
            else
            {
                switch (esMsgMarker.CurrentStatus)
                {
                    case MessageStatus.Ok:
                        result = new WPFReportCreator(wpfEntityColl, CommandCheck.ESEntity.CheckName, esMsgRun.Description, esMsgMarker.Description);
                        break;

                    case MessageStatus.Error:
                        if (isMarkered)
                        {
                            TaskDialog taskDialog = new TaskDialog("[ОШИБКА]")
                            {
                                MainInstruction = $"{CommandCheck.ESEntity.CheckName}: {esMsgMarker.Description}"
                            };
                            taskDialog.Show();
                            return null;
                        }
                        else
                            result = new WPFReportCreator(wpfEntityColl, CommandCheck.ESEntity.CheckName, esMsgRun.Description);

                        break;
                }
            }

            return result;
        }
    }
}

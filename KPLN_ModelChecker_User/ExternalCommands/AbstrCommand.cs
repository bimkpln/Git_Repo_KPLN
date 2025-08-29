using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    public abstract class AbstrCommand<T>
    {
        /// <summary>
        /// Конструктор для классов, наследуемых от AbstrCheckCommand. Если его не переопределить в наследнике - IExternalCommand не справиться с запуском (ему нужен конструтор по умолчанию)
        /// </summary>
        public AbstrCommand()
        {
        }

        /// <summary>
        /// Конструктор для класса Module. Он инициализирует основные переменные для работы с ExtensibleStorage
        /// </summary>
        internal AbstrCommand(ExtensibleStorageEntity esEntity)
        {
            ESEntity = esEntity;
        }

        /// <summary>
        /// Коллекция элементов с ошибками
        /// </summary>
        public static CheckerEntity[] CheckerEntities { get; private protected set; }

        /// <summary>
        /// Ссылка на ExtensibleStorageEntity
        /// </summary>
        internal static ExtensibleStorageEntity ESEntity { get; private protected set; }

        /// <summary>
        /// Ссылка на проверку
        /// </summary>
        public AbstrCheck<T> CommandCheck { get; set; }

        /// <summary>
        /// Спец. метод для вызова данного класса из кнопки WPF: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        public abstract Result ExecuteByUIApp(UIApplication uiapp, bool setPluginActivity = false, bool showMainForm = false, bool showSuccsessText = false);

        /// <summary>
        /// Подготовка окна результата проверки для пользователя
        /// </summary>
        /// <param name="uiapp">Revit-UIApplication</param>
        /// <param name="isElemZoom">Можно ли зумировать по элементу?</param>
        /// <param name="isMarkered">Нужно ли использовать основной маркер при создании окна?</param>
        /// <returns>Окно для вывода пользователю</returns>
        public void ReportCreatorAndDemonstrator(UIApplication uiapp, bool isMarkered = false)
        {
            WPFReportCreator repCreator = CreateReport(uiapp.ActiveUIDocument.Document, isMarkered);
            SetWPFEntityFiltration(repCreator);

            CheckMainForm form = new CheckMainForm(uiapp, this.GetType().Name, repCreator, ESEntity.ESBuilderRun, ESEntity.ESBuilderUserText, ESEntity.ESBuildergMarker);
            form.Show();
        }

        /// <summary>
        /// Установить фильтрацию элементов в отчете
        /// </summary>
        private protected abstract void SetWPFEntityFiltration(WPFReportCreator report);

        /// <summary>
        /// Метод для подготовки отчета
        /// </summary>
        private WPFReportCreator CreateReport(Document doc, bool isMarkered)
        {
            #region Настройка информации по логам проека
            Element piElem = doc.ProjectInformation;
            ResultMessage esMsgRun = ESEntity.ESBuilderRun.GetResMessage_Element(piElem);
            ResultMessage esMsgMarker = ESEntity.ESBuildergMarker.GetResMessage_Element(piElem);
            #endregion

            WPFReportCreator result = null;
            WPFEntity[] wpfEntityColl = CheckerEntities.Select(chEnt => new WPFEntity(chEnt, ESEntity)).ToArray();
            if (ESEntity.ESBuildergMarker.Guid.Equals(Guid.Empty))
                result = new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description);
            else
            {
                switch (esMsgMarker.CurrentStatus)
                {
                    case MessageStatus.Ok:
                        result = new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description, esMsgMarker.Description);
                        break;

                    case MessageStatus.Error:
                        if (isMarkered)
                        {
                            TaskDialog taskDialog = new TaskDialog("[ОШИБКА]")
                            {
                                MainInstruction = $"{ESEntity.CheckName}: {esMsgMarker.Description}"
                            };
                            taskDialog.Show();
                            return null;
                        }
                        else
                            result = new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description);

                        break;
                }
            }

            // Логируем последний запуск (отдельно, если все было ОК, а потом всплыли ошибки)
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESEntity.ESBuilderRun, DateTime.Now));

            return result;
        }
    }
}

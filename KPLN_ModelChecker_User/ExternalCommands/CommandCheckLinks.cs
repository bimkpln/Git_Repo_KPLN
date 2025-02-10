using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckLinks : AbstrCheckCommand<CommandCheckLinks>, IExternalCommand
    {
        internal const string PluginName = "Проверка связей";

        public CommandCheckLinks() : base()
        {
        }

        internal CommandCheckLinks(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            Element[] rvtLinks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .WhereElementIsNotElementType()
                // Фильтрация по имени от вложенных прикрепленных связей
                .Where(e => e.Name.Split(new string[] { ".rvt : " }, StringSplitOptions.None).Length < 3)
                .ToArray();

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, rvtLinks);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl)
        {
            if (!doc.IsWorkshared) throw new CheckerException("Проект не для совместной работы. Работа над такими проектами запрещена BEP");

            if (!(objColl.Any())) throw new CheckerException("В проекте отсутсвуют связи");

            foreach (object obj in objColl)
            {
                if (obj is Element element)
                {
                    if (element is RevitLinkInstance revitLink)
                    {
                        Document document = revitLink.GetLinkDocument() ?? throw new CheckerException($"Необходимо загрузить ВСЕ связи. Проверь диспетчер Revit-связей");
                    }
                    else
                        throw new Exception("Ошибка определения RevitLinkInstance");
                }
                else throw new Exception("Ошибка анализируемой коллекции");

            }

            return Enumerable.Empty<CheckCommandError>();
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            result.AddRange(CheckLocation(doc, elemColl));
            if (CheckPin(elemColl) is WPFEntity checkPin)
                result.Add(checkPin);

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Проверка на корректность общей площадки
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="rvtLinks">Коллекция связей</param>
        /// <returns>Коллекция ошибок WPFEntity</returns>
        private IEnumerable<WPFEntity> CheckLocation(Document doc, IEnumerable<Element> rvtLinks)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            IEnumerable<string> rvtLinkDocTitles = rvtLinks
                .Cast<RevitLinkInstance>()
                .Select(rli => rli.GetLinkDocument().Title);

            foreach (RevitLinkInstance link in rvtLinks.Cast<RevitLinkInstance>())
            {
                if (IsLinkWithSharedCoordByName(link))
                    continue;

                // Анализ наличия нескольких экземпляров связей
                if (rvtLinkDocTitles.Count(title => title == link.GetLinkDocument().Title) > 1)
                {
                    result.Add(new WPFEntity(
                        ESEntity,
                        link,
                        "Ошибка размещения",
                        "Экземпляры данной связи размещены несколько раз",
                        "Проверку необходимо выполнить вручную. Положение связей задается ТОЛЬКО через общую площадку, а наличие нескольких экземпляров разрешено только для типовых этажей."
                            + "\nВАЖНО: в отчет попала связь БЕЗ площадки, скорее всего её нужно удалить, но всё зависит от конкретного случая",
                        false));
                }

                // Анализ ОП в линках
                using (Transaction transaction = new Transaction(doc))
                {
                    try
                    {
                        transaction.Start("KPLN_CheckLink");
                        // Попытка получения координат из связи. Если общ. площадка отсутсвует - попытка будет успешной (т.е. ошибка, которая должна уйти в отчет),
                        // иначе - InvalidOperationException 
                        doc.AcquireCoordinates(link.Id);
                        result.Add(new WPFEntity(
                            ESEntity,
                            link,
                            "Ошибка размещения",
                            "У связи и проекта - разные системы координат",
                            "Запрещено размещать связи без общих площадок",
                            false));
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
                    {
                        if (ioe.Message.Contains("The coordinate system of the selected model are the same as the host model"))
                        {
                            result.Add(new WPFEntity(
                                ESEntity,
                                link,
                                "Ошибка размещения",
                                "У связи не выбрана общая площадка",
                                "Запрещено размещать связи без общих площадок",
                                false));
                        }
                        else if (ioe.Message.Contains("Failed to acquire coordinates from the link instance"))
                        {
                            result.Add(new WPFEntity(
                                ESEntity,
                                link,
                                "Ошибка размещения",
                                "У связи не удалось получить координаты",
                                "Может быть связано с внутренними проблемами, например - занят рабочий набор \"Сведения о проекте\"",
                                false));
                        }
                        else
                            throw new Exception($"Ошибка проверки связей: {ioe.Message} для файла {link.Name}");
                    }
                    finally
                    {
                        transaction.RollBack();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Проверка экз. связи на наличие общей площадки из имени. ВАЖНО: Слабоватый метод, т.к. имена со временем могут поменяться, но на текущий момент - другого не вижу
        /// </summary>
        /// <param name="rli"></param>
        /// <returns></returns>
        [Obsolete]
        private bool IsLinkWithSharedCoordByName(RevitLinkInstance rli) =>
           !(rli.Name.ToLower().Contains("<not shared>") || rli.Name.ToLower().Contains("не общедоступное"));

        /// <summary>
        /// Проверка закрепление (PIN)
        /// </summary>
        /// <param name="rvtLinks">Коллекция связей</param>
        /// <returns>Коллекция ошибок WPFEntity</returns>
        private WPFEntity CheckPin(IEnumerable<Element> rvtLinks)
        {
            List<Element> errorElems = new List<Element>();
            foreach (RevitLinkInstance link in rvtLinks)
            {
                if (!link.Pinned) errorElems.Add(link);
            }

            if (errorElems.Any())
            {
                return new WPFEntity(
                    ESEntity,
                    errorElems,
                    "Ошибка прикрепления",
                    "Связи необходимо прикрепить (команда 'Прикрепить' ('Pin')) ВНИМАНИЕ: не путать с настройкой типа связи 'Прикрепление' ('Attachment')",
                    "Может быть связано с внутренними проблемами, например - занят рабочий набор \"Сведения о проекте\"",
                    false);
            }

            return null;
        }
    }
}
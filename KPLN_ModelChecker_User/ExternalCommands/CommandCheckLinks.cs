using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckLinks : AbstrCheckCommand<CommandCheckLinks>, IExternalCommand
    {
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
                        Document document = revitLink.GetLinkDocument();
                        if (document == null) throw new CheckerException($"Необходимо загрузить ВСЕ связи. Проверь диспетчер Revit-связей");
                    }
                    else throw new Exception("Ошибка определения RevitLinkInstance");
                }
                else throw new Exception("Ошибка анализируемой коллекции");

            }

            return Enumerable.Empty<CheckCommandError>();
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            result.AddRange(CheckLocation(doc, elemColl));
            if (CheckPin(elemColl) is WPFEntity checkPin) result.Add(checkPin);

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

            // Анализ ОП в линках
            foreach (RevitLinkInstance link in rvtLinks)
            {
                using (Transaction transaction = new Transaction(doc))
                {
                    try
                    {
                        transaction.Start("KPLN_CheckLink");
                        // Попытка получения координат из связи. Если общ. площадка отсутсвует - попытка будет успешной (т.е. ошибка, которая должна уйти в отчет),
                        // иначе - InvalidOperationException 
                        doc.AcquireCoordinates(link.Id);
                        result.Add(new WPFEntity(
                            link,
                            CheckStatus.Error,
                            "Ошибка размещения",
                            "У связи и проекта - разные системы координат",
                            false,
                            false,
                            "Запрещено размещать связи без общих площадок"));
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
                    {
                        if (ioe.Message.Contains("The coordinate system of the selected model are the same as the host model."))
                        {
                            // Слабоватый метод, т.к. имена со временем могут поменяться, но на текущий момент - другого не вижу
                            if (link.Name.ToLower().Contains("<not shared>") || link.Name.ToLower().Contains("не общедоступное"))
                            {
                                result.Add(new WPFEntity(
                                    link,
                                    CheckStatus.Error,
                                    "Ошибка размещения",
                                    "У связи не выбрана общая площадка",
                                    false,
                                    false,
                                    "Запрещено размещать связи без общих площадок"));
                            }
                            else
                                continue;
                        }
                        else if (ioe.Message.Contains("Cannot acquire coordinates from a model placed multiple times."))
                        {
                            result.Add(new WPFEntity(
                                link,
                                CheckStatus.Warning,
                                "Ошибка размещения",
                                "Экземпляры данной связи размещены несколько раз",
                                false,
                                false,
                                "Проверку необходимо выполнить вручную. Положение связей задается ТОЛЬКО через общую площадку"));
                        }
                        else 
                            throw new Exception($"Ошибка проверки связей: {ioe.Message}");
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
                    errorElems,
                    CheckStatus.Error,
                    "Ошибка прикрепления",
                    "Связи необходимо прикрепить (команда 'Прикрепить' ('Pin')) ВНИМАНИЕ: не путать с настройкой типа связи 'Прикрепление' ('Attachment')",
                    false,
                    false);
            }

            return null;
        }
    }
}
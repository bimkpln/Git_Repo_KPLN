using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckLinks : AbstrCheckCommand, IExternalCommand
    {
        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            _name = "Проверка связей";
            _application = uiapp;

            _allStorageName = "KPLN_CheckLinks";

            _lastRunGuid = new Guid("045e7890-0ff3-4be3-8f06-1fa1dd7e762e");
            _userTextGuid = new Guid("045e7890-0ff3-4be3-8f06-1fa1dd7e762f");

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            IEnumerable<Element> rvtLinks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .WhereElementIsNotElementType()
                // Фильтрация по имени от вложенных прикрепленных связей
                .Where(e => e.Name.Split(new string[] { ".rvt : " }, StringSplitOptions.None).Length < 3);

            #region Проверяю и обрабатываю элементы
            IEnumerable<WPFEntity> wpfColl = CheckCommandRunner(doc, rvtLinks);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override List<CheckCommandError> CheckElements(Document doc, IEnumerable<Element> elemColl)
        {
            if (!doc.IsWorkshared) throw new UserException("Проект не для совместной работы. Работа над такими проектами запрещена BEP");

            if (!(elemColl.Any())) throw new UserException("В проекте отсутсвуют связи");

            foreach (Element element in elemColl)
            {
                if (element is RevitLinkInstance revitLink)
                {
                    Document document = revitLink.GetLinkDocument();
                    if (document == null) throw new UserException($"Необходимо загрузить ВСЕ связи. Проверь диспетчер Revit-связей");
                }
                else throw new Exception("Ошибка определения RevitLinkInstance");
            }

            return null;
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, IEnumerable<Element> elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            result.AddRange(CheckWorkSets(elemColl));
            result.AddRange(CheckLocation(doc, elemColl));
            if (CheckPin(elemColl) is WPFEntity checkPin) result.Add(checkPin);

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Проверка на корректность рабочих наборов
        /// </summary>
        /// <param name="rvtLinks">Коллекция связей</param>
        /// <returns>Коллекция ошибок WPFEntity</returns>
        private IEnumerable<WPFEntity> CheckWorkSets(IEnumerable<Element> rvtLinks)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (RevitLinkInstance link in rvtLinks)
            {
                string[] separators = { ".rvt : " };
                string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                if (nameSubs.Length > 3) continue;

                string wsName = link.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                if (!wsName.StartsWith("00") && !wsName.StartsWith("#"))
                {
                    result.Add(new WPFEntity(
                        link,
                        Status.Error,
                        "Ошибка рабочего набора",
                        "Связь находится в некорректном рабочем наборе",
                        false,
                        false,
                        "Для связей необходимо использовать именные рабочие наборы, которые начинаются с '00_'"));
                }
            }

            return result;
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
                            Status.Error,
                            "Ошибка размещения",
                            "У связи не указана общая площадка",
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
                                    Status.Error,
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
                                Status.Warning,
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
                    Status.Error,
                    "Ошибка прикрепления",
                    "Связи необходимо прикрепить (команда 'Прикрепить' ('Pin'). Не путать с настройкой типа связи 'Прикрепление' ('Attachment'))",
                    false,
                    false);
            }

            return null;
        }
    }
}
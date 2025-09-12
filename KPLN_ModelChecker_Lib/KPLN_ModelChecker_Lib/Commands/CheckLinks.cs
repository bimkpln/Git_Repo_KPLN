using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckLinks : AbstrCheck
    {
        /// <summary>
        /// Пустой конструктор для внесения данных класса
        /// </summary>
        public CheckLinks() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка связей";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(PluginName, "KPLN_CheckLinks", new Guid("045e7890-0ff3-4be3-8f06-1fa1dd7e762e"));
        }


        public override Element[] GetElemsToCheck(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .WhereElementIsNotElementType()
                // Фильтрация по имени от вложенных прикрепленных связей
                .Where(e => e.Name.Split(new string[] { ".rvt : " }, StringSplitOptions.None).Length < 3)
                .ToArray();
        }

        private protected override IEnumerable<CheckCommandError> CheckRElems(object[] objColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<CheckerEntity> GetCheckerEntities(Document doc, Element[] elemColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            RevitLinkInstance[] openedRLIColl = elemColl
                .Cast<RevitLinkInstance>()
                .Where(rli => rli.GetLinkDocument() != null)
                .ToArray();

            result.AddRange(CheckLocation(openedRLIColl));

            CheckerEntity checkPIN = CheckPin(openedRLIColl);
            if (checkPIN != null)
                result.Add(checkPIN);

            result.AddRange(CheckPath(openedRLIColl));

            return result;
        }

        /// <summary>
        /// Проверка на корректность общей площадки
        /// </summary>
        /// <param name="rliColl">Коллекция связей</param>
        /// <returns>Коллекция ошибок WPFEntity</returns>
        private IEnumerable<CheckerEntity> CheckLocation(IEnumerable<RevitLinkInstance> rliColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            IEnumerable<string> rvtLinkDocTitles = rliColl.Select(rli => rli.GetLinkDocument().Title);

            foreach (RevitLinkInstance link in rliColl)
            {
                if (IsLinkWithSharedCoordByNameError(link))
                {
                    result.Add(new CheckerEntity(
                        link,
                        "Ошибка общей площадки",
                        "У связи не выбрана общая площадка",
                        "Запрещено размещать связи без общих площадок, т.к. может быть ошибка в пространственном положении связи. " +
                            "Если не знаешь как исправить - обратись в BIM-отдел",
                        false));
                }

                // Анализ наличия нескольких экземпляров связей
                if (rvtLinkDocTitles.Count(title => title == link.GetLinkDocument().Title) > 1)
                {
                    result.Add(new CheckerEntity(
                        link,
                        "Ошибка размещения",
                        "Экземпляры данной связи размещены несколько раз",
                        "Проверку необходимо выполнить вручную. Положение связей задается ТОЛЬКО через общую площадку, а наличие нескольких экземпляров разрешено только для типовых этажей."
                            + "\nВАЖНО: в отчет попала связь БЕЗ площадки, скорее всего её нужно удалить, но всё зависит от конкретного случая",
                        false));
                }

                // Анализ ОП в линках (выкинул в архив, т.к. для автопроверок постоянно занимается РН, плюс юзеру важен факт - дальше сам разберёться)
                //using (Transaction transaction = new Transaction(CheckDoc))
                //{
                //    try
                //    {
                //        transaction.Start("KPLN_CheckLink");
                //        // Попытка получения координат из связи. Если общ. площадка отсутсвует - попытка будет успешной (т.е. ошибка, которая должна уйти в отчет),
                //        // иначе - InvalidOperationException 
                //        CheckDoc.AcquireCoordinates(link.Id);
                //        result.Add(new CheckerEntity(
                //            link,
                //            "Ошибка размещения",
                //            "У связи и проекта - разные системы координат",
                //            "Запрещено размещать связи без общих площадок, т.к. может быть ошибка в пространственном положении связи",
                //            false));
                //    }
                //    catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
                //    {
                //        if (ioe.Message.Contains("The coordinate system of the selected model are the same as the host model"))
                //        {
                //            result.Add(new CheckerEntity(
                //                link,
                //                "Ошибка размещения",
                //                "У связи не выбрана общая площадка",
                //                "Запрещено размещать связи без общих площадок, т.к. может быть ошибка в пространственном положении связи",
                //                false));
                //        }
                //        else if (ioe.Message.Contains("Failed to acquire coordinates from the link instance"))
                //        {
                //            result.Add(new CheckerEntity(
                //                link,
                //                "Ошибка размещения",
                //                "У связи не удалось получить координаты",
                //                "Может быть связано с внутренними проблемами, например - занят рабочий набор \"Сведения о проекте\"",
                //                false));
                //        }
                //        else
                //            throw new Exception($"Ошибка проверки связей: {ioe.Message} для файла {link.Name}");
                //    }
                //    finally
                //    {
                //        transaction.RollBack();
                //    }
                //}
            }
            return result;
        }

        /// <summary>
        /// Проверка экз. связи на наличие общей площадки из имени
        /// </summary>
        private bool IsLinkWithSharedCoordByNameError(RevitLinkInstance rli) => rli.Name.ToLower().Contains("<not shared>") || rli.Name.ToLower().Contains("не общедоступное");

        /// <summary>
        /// Проверка закрепление (PIN)
        /// </summary>
        /// <param name="rliColl">Коллекция связей</param>
        /// <returns>Коллекция ошибок WPFEntity</returns>
        private CheckerEntity CheckPin(IEnumerable<RevitLinkInstance> rliColl)
        {
            List<Element> errorElems = new List<Element>();
            foreach (RevitLinkInstance link in rliColl)
            {
                if (!link.Pinned) errorElems.Add(link);
            }

            if (errorElems.Any())
            {
                return new CheckerEntity(
                    errorElems,
                    "Ошибка прикрепления",
                    "Связи необходимо прикрепить (команда 'Прикрепить' ('Pin')) ВНИМАНИЕ: не путать с настройкой типа связи 'Прикрепление' ('Attachment')",
                    "Это позволит избежать случайного смещения связи пользователем, при работе с моделью",
                    false);
            }

            return null;
        }

        /// <summary>
        /// Проверка пути линка (откуда загружен)
        /// </summary>
        /// <param name="rliColl">Коллекция связей</param>
        /// <returns></returns>
        private IEnumerable<CheckerEntity> CheckPath(IEnumerable<RevitLinkInstance> rliColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();
            foreach (RevitLinkInstance link in rliColl)
            {
                string lDocPath;
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc.IsWorkshared)
                {
                    ModelPath lDocMPath = linkDoc.GetWorksharingCentralModelPath();
                    lDocPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(lDocMPath);
                }
                else
                    lDocPath = linkDoc.PathName;

                if (lDocPath.ToLower().Contains("архив"))
                    result.Add(new CheckerEntity(
                        link,
                        "Подозрительный путь",
                        $"Большая вероятность, что связь случайно выбрана из архива, т.к. путь: {lDocPath}",
                        $"Ошибка может быть ложной, если на данном проекте принято такое решение. Перед заменой - ПРОКОНСУЛЬТИРУЙСЯ в BIM-отделе",
                        false,
                        ErrorStatus.Warning));
            }

            return result;
        }
    }
}

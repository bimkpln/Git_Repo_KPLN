using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_IOSClasher.Core;
using KPLN_IOSClasher.ExecutableCommand;
using KPLN_IOSClasher.Services;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_IOSClasher
{
    public class Module : IExternalModule
    {
        /// <summary>
        /// Общий фильтр для просеивания элементов модели (новых и отредактированных)
        /// </summary>
        private static Func<Element, bool> _elemFilterFunc;

        public Module()
        {
            ModuleDBWorkerService = new DBWorkerService();

            _elemFilterFunc = (el) => 
                el.Category != null
                && !(el is ElementType)
                && IntersectCheckEntity.BuiltInCatIDs.Any(bicId => el.Category.Id.IntegerValue == bicId)
                // Игнор огнезащиты для ЭОМСС, которые моделируются воздуховодами
                && !(el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("ASML_ОГК_")
                    || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("Огнезащитный короб_EI150"));
        }

        public static DBWorkerService ModuleDBWorkerService { get; private set; }

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
#if Revit2020 || Revit2023
            //Фильтрация по разделам
            if (ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("АР")
                || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("КР")
                || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                return Result.Succeeded;
#endif
            // Персональная фильтрация по сотрудникам (далее повесить на БД, пока хардкод)
            if (ModuleDBWorkerService.CurrentDBUser.Id == 172
                || ModuleDBWorkerService.CurrentDBUser.Id == 71
                || ModuleDBWorkerService.CurrentDBUser.Id == 111
                || ModuleDBWorkerService.CurrentDBUser.Id == 126)
                return Result.Succeeded;


            //Подписка на события
            application.ViewActivated += OnViewActivated;
            application.ControlledApplication.DocumentChanged += OnDocumentChanged;
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;

            return Result.Succeeded;
        }

        /// <summary>
        /// Событие на открытый документ
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            if (doc.IsFamilyDocument)
                return;

#if Revit2020 || Revit2023
            // Игнор НЕ мониторинговых моделей
            if (KPLN_Looker.Module.MonitoredDocFilePath(doc) == null)
                return;
#endif
            DocumentSet docSet = doc.Application.Documents;
            bool haveLinks = false;
            foreach (Document openDoc in docSet)
            {
                if (openDoc.IsLinked)
                {
                    haveLinks = true;
                    break;
                }
            }

            if (haveLinks)
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new IntersectPointLinkWorker());
        }

        /// <summary>
        /// Событие на активацию вида
        /// </summary>
        private static void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            Document doc = args.Document;
            if (doc.IsFamilyDocument)
                return;

#if Revit2020 || Revit2023
            // Игнор НЕ мониторинговых моделей
            if (KPLN_Looker.Module.MonitoredDocFilePath(doc) == null)
                return;
#endif
            View activeView = args.CurrentActiveView;

            #region Обновление кэша по сервисам
            ViewType actViewType = activeView.ViewType;
            // Если вид НЕ модельный - игнор
            if (actViewType == ViewType.FloorPlan
                || actViewType == ViewType.ThreeD
                || actViewType == ViewType.CeilingPlan
                || actViewType == ViewType.Section
                || actViewType == ViewType.EngineeringPlan
                || actViewType == ViewType.Elevation)
            {
                DocController.CurrentDocumentUpdateData(doc);
#if Revit2020 || Revit2023
                // Если не анализируется, то и линки не трогаю
                if (!DocController.IsDocumentAnalyzing)
                    return;
#endif

                DocController.UpdateIntCheckEntities_Link(doc, activeView);
            }
            #endregion
        }

        /// <summary>
        /// Событие на изменение в документе
        /// </summary>
        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            Document doc = args.GetDocument();
            if (doc.IsFamilyDocument)
                return;

            // Игнор НЕ мониторинговых моделей
#if Revit2020 || Revit2023
            if (!DocController.IsDocumentAnalyzing)
                return;
#endif

            string transName = args.GetTransactionNames().FirstOrDefault();
            // Обновляю по линкам, если были транзакции
            if (// Рунчая загрузка связи
                transName.Equals("Связать с проектом Revit")
                // Рунчая загрузка связи
                || transName.Equals("Загрузить связь")
                // Выгрузить линк для всех
                || transName.Equals("Выгрузить связь")
                // Выгрузить линк для меня
                || transName.Equals("Для меня")
                // Открыть РН
                || transName.Equals("Повторная загрузка")
                // Закрыть РН
                || transName.Equals("Выгрузить связь")
                // Перетаскивание границы подрезки (главное, чтобы не листах, т.к. там такая же транзакция на перенос границ столбцов спецификаций, текста и пр.)
                || (transName.Equals("Перенести") && doc.ActiveView != null && doc.ActiveView.ViewType != ViewType.DrawingSheet)
                // Редактирование границы подрезки
                || transName.Equals("Принять эскиз"))
            {
                UIDocument uidoc = new UIDocument(doc);
                View activeView = uidoc.ActiveView ?? throw new Exception("Отправь разработчику - не удалось определить класс View");

                // Анализирую коллизии по линкам
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new IntersectPointLinkWorker());

                // Обновляю коллекцию на потенциальных элементов ЛИНКОВ по триггерам обновления линков
                DocController.UpdateIntCheckEntities_Link(doc, activeView);
            }

            Element[] addedLinearElems = args
                .GetAddedElementIds()
                .Select(id => doc.GetElement(id))
                .Where(_elemFilterFunc)
                .ToArray();

            Element[] modifyedLinearElems = args
                .GetModifiedElementIds()
                .Select(id => doc.GetElement(id))
                .Where(_elemFilterFunc)
                .ToArray();

            ElementId[] deletedLinearElems = args
                .GetDeletedElementIds()
                .ToArray();

            if (deletedLinearElems.Any())
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new IntersectPointCleaner(deletedLinearElems));

            Element[] allChangedElems = addedLinearElems.Concat(modifyedLinearElems).ToArray();

            if (!allChangedElems.Any())
                return;

            // Обновляю коллекцию на потенциальных элементов ВНУТРИ документа
            DocController.UpdateIntCheckEntities_Doc(doc, allChangedElems);

            HashSet<IntersectPointEntity> intersectedPointEntities = new HashSet<IntersectPointEntity>();
            // Тут и далее хардкод по разделам КПЛН для экономии ресурсов
            switch (DocController.CheckDocDBSubDepartmentId)
            {
                case -1:
                    goto case 99;
                case 1:
                    goto case 99;
                case 2:
                    goto case 99;
                case 3:
                    goto case 99;
                case 4:
                    DocController.CheckIf_OVLoad = true;
                    goto case 98;
                case 5:
                    DocController.CheckIf_VKLoad = true;
                    goto case 98;
                case 6:
                    DocController.CheckIf_EOMLoad = true;
                    goto case 98;
                case 7:
                    DocController.CheckIf_SSLoad = true;
                    goto case 98;
                case 98:
                    intersectedPointEntities = DocController.GetIntPntEntities(allChangedElems);
                    break;
                case 99:
                    return;
            }

            bool isAllLinkLoad = DocController.CheckIf_OVLoad
                && DocController.CheckIf_VKLoad
                && DocController.CheckIf_EOMLoad &&
                DocController.CheckIf_SSLoad;

            // ВСЕ ИОС связи должны быть загружены
            if (!(isAllLinkLoad))
            {
                if (!KPLN_Loader.Application.OnIdling_CommandQueue.Any(i => i is LinkAlarmShower))
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new LinkAlarmShower());
            }

            if (allChangedElems.Any())
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new IntersectPointMaker(intersectedPointEntities, allChangedElems));
        }
    }
}

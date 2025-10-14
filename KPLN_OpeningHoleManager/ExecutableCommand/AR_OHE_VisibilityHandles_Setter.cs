using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по расширению ручек у семейств АР
    /// </summary>
    internal sealed class AR_OHE_VisibilityHandles_Setter : IExecutableCommand
    {
        /// <summary>
        /// Коллекция для группирования по сообщению и вывода данных (Id-элементов) пользователю
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _msgDict_ByMsg = new Dictionary<string, List<ElementId>>();
        private readonly IEnumerable<AROpeningHoleEntity> _arEntities;

        private readonly ProgressInfoViewModel _progressInfoViewModel;

        public AR_OHE_VisibilityHandles_Setter(IEnumerable<AROpeningHoleEntity> arEntities, ProgressInfoViewModel progressInfoViewModel)
        {
            _arEntities = arEntities;
            _progressInfoViewModel = progressInfoViewModel;
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;

            Module.CurrentUIContrApp.ControlledApplication.FailuresProcessing += RevitEventWorker.OnFailureProcessing;

            try
            {
                using (Transaction trans = new Transaction(doc, "KPLN: Расширить ручки видимости"))
                {
                    trans.Start();

                    _progressInfoViewModel.CurrentProgress = 0;
                    _progressInfoViewModel.MaxProgress = _arEntities.Count();
                    _progressInfoViewModel.ProcessTitle = $"Расширение границ видимости...";

                    foreach (AROpeningHoleEntity arEnt in _arEntities)
                    {
                        ParamWriter(arEnt);

                        ++_progressInfoViewModel.CurrentProgress;
                        _progressInfoViewModel.DoEvents();
                    }
                    _progressInfoViewModel.IsComplete = true;

                    trans.Commit();
                }
                HtmlOutput.PrintMsgDict("Внимание", MessageType.Warning, _msgDict_ByMsg);
            }
            catch (CheckerException ex)
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                    MainInstruction = $"{ex.Message}",
                }.Show();

                return Result.Cancelled;
            }
            catch (Exception ex) { throw ex; }
            finally
            {
                Module.CurrentUIContrApp.ControlledApplication.FailuresProcessing -= RevitEventWorker.OnFailureProcessing;
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Определить и заполнить параметры
        /// </summary>
        private void ParamWriter(AROpeningHoleEntity arEnt)
        {
            Parameter upExpandParam = arEnt.IEDElem.LookupParameter(arEnt.AR_OHE_VisibilityHandles_ParamNameUpExpander);
            Parameter downExpandParam = arEnt.IEDElem.LookupParameter(arEnt.AR_OHE_VisibilityHandles_ParamNameDownExpander);
            if (upExpandParam == null || downExpandParam == null)
                throw new CheckerException($"У элемента с id: {arEnt.IEDElem.Id} нет одного из параметров для расширения границ. Обратись к разработчику");

            // СЕТ: Поправка, если отверстия окном (окна смещаются по пар-ру, привязка у солида остаётся снизу).
            // "Высота нижнего бруса"
            double sillHeight = 0;
            if (arEnt.IEDElem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
            {
                if (arEnt.OHE_Shape == Core.MainEntity.OpenigHoleShape.Rectangular)
                    sillHeight = arEnt.IEDElem.LookupParameter("АР_Высота Проема").AsDouble();
            }

            double upLvlResultDistance = 1.64042;
            if (arEnt.UpFloorDistance < double.MaxValue)
                upLvlResultDistance = arEnt.UpFloorDistance;
            else
                HtmlOutput.SetMsgDict_ByMsg(
                    "У элемента не определен уровень сверху. Установлено значение по умолчанию (500 мм). Убедись, что все связи АР загружены, т.к. анализ идёт на основе перекрытий АР",
                    arEnt.IEDElem.Id, 
                    _msgDict_ByMsg);

            double downLvlResultDistance = 1.64042;
            if (arEnt.DownFloorDistance < double.MaxValue)
                downLvlResultDistance = arEnt.DownFloorDistance;
            else
                HtmlOutput.SetMsgDict_ByMsg(
                    "У элемента не определен уровень снизу. Установлено значение по умолчанию (500 мм). Убедись, что все связи АР загружены, т.к. анализ идёт на основе перекрытий АР",
                    arEnt.IEDElem.Id,
                    _msgDict_ByMsg);


            upExpandParam.Set(upLvlResultDistance);
            downExpandParam.Set(Math.Abs(downLvlResultDistance + sillHeight));
        }
    }
}

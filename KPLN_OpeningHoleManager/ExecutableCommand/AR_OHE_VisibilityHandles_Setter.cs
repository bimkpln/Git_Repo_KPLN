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
        private readonly List<AROpeningHoleEntity> _warningArEntities_UpLvl = new List<AROpeningHoleEntity>();
        private readonly List<AROpeningHoleEntity> _warningArEntities_DownLvl = new List<AROpeningHoleEntity>();

        private const double _defaultExpandValue = 1.64042;
        
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

                    PrepareProgress("Расширение границ видимости...", 0, _arEntities.Count());

                    foreach (AROpeningHoleEntity arEnt in _arEntities)
                    {
                        Parameter upParam = GetParam(arEnt, arEnt.AR_OHE_VisibilityHandles_ParamNameUpExpander);
                        Parameter downParam = GetParam(arEnt, arEnt.AR_OHE_VisibilityHandles_ParamNameDownExpander);

                        double sillHeight = GetSillHeight(arEnt);

                        SetParamValue(arEnt, upParam, arEnt.UpFloorDistance + sillHeight, _warningArEntities_UpLvl, "сверху");
                        SetParamValue(arEnt, downParam, arEnt.DownFloorDistance, _warningArEntities_DownLvl, "снизу");

                        UpdateProgress();
                    }

                    if (_msgDict_ByMsg.Any())
                        ProcessWarningsWithUserChoice();

                    trans.Commit();
                    
                    _progressInfoViewModel.IsComplete = true;
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
        /// Выбор пользователем при наличии замечаний
        /// </summary>
        private void ProcessWarningsWithUserChoice()
        {
            int warnCount = _warningArEntities_UpLvl.Count + _warningArEntities_DownLvl.Count;
            TaskDialog td = new TaskDialog("ВНИМАНИЕ: Закройте вид!")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainInstruction = $"Выявлены элементы ({warnCount} шт.), без оснований. Записать им значение по умолчанию (500 мм)?\n" +
                                  $"ИНФО: Поиск оснований идёт по плитам АР, включая связи.",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
            };

            if (td.Show() == TaskDialogResult.Yes)
            {
                PrepareProgress("Расширение границ видимости (значение по умолчанию)...", _arEntities.Count() - warnCount, _arEntities.Count());

                ApplyDefaultValues(_warningArEntities_UpLvl, e => e.AR_OHE_VisibilityHandles_ParamNameUpExpander);
                ApplyDefaultValues(_warningArEntities_DownLvl, e => e.AR_OHE_VisibilityHandles_ParamNameDownExpander);
            }
        }

        /// <summary>
        /// Установить значения по умолчанию для нужного параметра
        /// </summary>
        private void ApplyDefaultValues(IEnumerable<AROpeningHoleEntity> entities, Func<AROpeningHoleEntity, string> paramNameSelector)
        {
            foreach (var arEnt in entities)
            {
                Parameter param = GetParam(arEnt, paramNameSelector(arEnt));
                param.Set(_defaultExpandValue);
                UpdateProgress();
            }
        }

        /// <summary>
        /// Получить параметр по имени
        /// </summary>
        private Parameter GetParam(AROpeningHoleEntity arEnt, string paramName)
        {
            return arEnt.IEDElem.LookupParameter(paramName)
                ?? throw new CheckerException($"У элемента с id: {arEnt.IEDElem.Id} отсутствует параметр \"{paramName}\"");
        }

        /// <summary>
        /// Поправка для отверстий (окна смещаются по пар - ру, привязка у солида остаётся снизу), например для СЕТ
        /// </summary>
        /// <returns></returns>
        private double GetSillHeight(AROpeningHoleEntity arEnt)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            bool isWindCat = arEnt.IEDElem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows;
#else
            bool isWindCat = arEnt.IEDElem.Category.BuiltInCategory == BuiltInCategory.OST_Windows;
#endif
            if (isWindCat && arEnt.OHE_Shape == Core.MainEntity.OpenigHoleShape.Rectangular)
                return arEnt.IEDElem.LookupParameter("АР_Высота Проема")?.AsDouble() ?? 0;

            return 0;
        }

        private void SetParamValue(
            AROpeningHoleEntity arEnt,
            Parameter param,
            double value,
            List<AROpeningHoleEntity> warnList,
            string position)
        {
            if (value < double.MaxValue)
                param.Set(value);
            else
            {
                warnList.Add(arEnt);
                SetWarning(arEnt, position);
            }
        }

        private void SetWarning(AROpeningHoleEntity arEnt, string levelPosition)
        {
            HtmlOutput.SetMsgDict_ByMsg(
                $"У элемента не определен уровень {levelPosition}. Установлено значение по умолчанию (500 мм). " +
                $"Убедись, что все связи АР загружены.",
                arEnt.IEDElem.Id,
                _msgDict_ByMsg);
        }

        private void PrepareProgress(string title, int current, int max)
        {
            _progressInfoViewModel.CurrentProgress = current;
            _progressInfoViewModel.MaxProgress = max;
            _progressInfoViewModel.ProcessTitle = title;
        }

        private void UpdateProgress()
        {
            _progressInfoViewModel.CurrentProgress++;
            _progressInfoViewModel.DoEvents();
        }
    }
}

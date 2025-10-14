using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Linq;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по расстановке отверстий АР
    /// </summary>
    internal sealed class AR_OHE_Maker : IExecutableCommand
    {
        private readonly string _transName;
        private readonly AROpeningHoleEntity[] _arEntitiesToCreate;

        private readonly ProgressInfoViewModel _progressInfoViewModel;

        /// <summary>
        /// Конструктор для обработки исходной коллекции, чтобы далее использовать полученный результат
        /// </summary>
        /// <param name="arEntitiesToCreate">Коллекция для создания</param>
        public AR_OHE_Maker(AROpeningHoleEntity[] arEntitiesToCreate, string transName, ProgressInfoViewModel progressInfoViewModel)
        {
            _arEntitiesToCreate = arEntitiesToCreate;
            _transName = transName;

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
                // 3d- вид для изоляция основы для отверстий
                View3D view = null;
                using (Transaction trans = new Transaction(doc, _transName))
                {
                    trans.Start();

                    _progressInfoViewModel.CurrentProgress = 0;
                    _progressInfoViewModel.MaxProgress = _arEntitiesToCreate.Length;
                    _progressInfoViewModel.ProcessTitle = $"Создание одиночных отверстий...";

                    // Работа с элемнтами
                    foreach (AROpeningHoleEntity arEntity in _arEntitiesToCreate)
                    {
                        arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc);

                        ++_progressInfoViewModel.CurrentProgress;
                        _progressInfoViewModel.DoEvents();
                    }
                    AROpeningHoleEntity.RegenerateDocAndSetSolids(doc, _arEntitiesToCreate);
                    _progressInfoViewModel.IsComplete = true;

                    // Работа с видом и выделением эл-в
                    view = ViewZoomCreator.SpecialViewCreator(
                        app,
                        _arEntitiesToCreate.Select(ent => ent.AR_OHE_HostElement).ToHashSet(new ElementComparerById()),
                        true);

                    app.ActiveUIDocument.Selection.SetElementIds(
                        _arEntitiesToCreate
                        .Where(ent => doc.GetElement(ent.IEDElem.Id).IsValidObject)
                        .Select(ent => ent.IEDElem.Id)
                        .ToArray());

                    trans.Commit();
                }

                ViewZoomCreator.SpecialViewOpener(app, view);
            }
            catch (Exception ex) { throw ex; }
            finally
            {
                Module.CurrentUIContrApp.ControlledApplication.FailuresProcessing -= RevitEventWorker.OnFailureProcessing;
            }

            return Result.Succeeded;
        }
    }
}

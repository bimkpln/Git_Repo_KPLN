using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по расстановке отверстия АР
    /// </summary>
    internal sealed class AR_OHE_Maker : IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Отверстие по заданию";

        private readonly AROpeningHoleEntity[] _arEntities;

        /// <summary>
        /// Конструктор для обработки исходной коллекции, чтобы далее использовать полученный результат
        /// </summary>
        /// <param name="arEntities"></param>
        public AR_OHE_Maker(ref AROpeningHoleEntity[] arEntities)
        {
            _arEntities = arEntities;
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                foreach (AROpeningHoleEntity arEntity in _arEntities)
                {
                    arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc, arEntity.AR_OHE_HostElement);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}

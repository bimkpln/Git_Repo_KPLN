using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по удалению отверстий из модели
    /// </summary>
    internal sealed class AR_OHE_ElementDeleter : IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Очистка отверстий";

        private readonly List<AROpeningHoleEntity> _arEntities;

        public AR_OHE_ElementDeleter(List<AROpeningHoleEntity> arEntities)
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
                    doc.Delete(arEntity.OHE_Element.Id);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}

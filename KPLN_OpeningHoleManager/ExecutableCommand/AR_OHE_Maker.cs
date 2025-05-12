using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    /// <summary>
    /// Класс по расстановке и чистке отверстий АР
    /// </summary>
    internal sealed class AR_OHE_Maker : IExecutableCommand
    {
        private readonly string _transName;
        private readonly AROpeningHoleEntity[] _arEntitiesToCreate;
        private AROpeningHoleEntity[] _arEntitiesToDelete;

        /// <summary>
        /// Конструктор для обработки исходной коллекции, чтобы далее использовать полученный результат
        /// </summary>
        /// <param name="arEntitiesToCreate">Коллекция для создания</param>
        public AR_OHE_Maker(AROpeningHoleEntity[] arEntitiesToCreate, string transName)
        {
            _arEntitiesToCreate = arEntitiesToCreate;
            _transName = transName;
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;


            using (Transaction trans = new Transaction(doc, _transName))
            {
                trans.Start();

                // Сначала нахожу и удаляю полностью поглащенные отверстия
                _arEntitiesToDelete = AROpeningHoleEntity.GetEntitesToDel_ByFullIntescect(doc, _arEntitiesToCreate);
                foreach (AROpeningHoleEntity arEntity in _arEntitiesToDelete)
                {
                    doc.Delete(arEntity.OHE_Element.Id);
                }
                
                // Затем создаю новые элементы (порядок важен, чтобы не создавать доп фильтрацию при поиске полностью поглащенных)
                foreach (AROpeningHoleEntity arEntity in _arEntitiesToCreate)
                {
                    arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc, arEntity.AR_OHE_HostElement);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}

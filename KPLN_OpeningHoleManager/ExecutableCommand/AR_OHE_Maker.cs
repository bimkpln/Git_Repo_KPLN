using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Services;
using System.Collections.Generic;
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


            // 3d- вид для изоляция основы для отверстий
            View3D view = null;
            using (Transaction trans = new Transaction(doc, _transName))
            {
                trans.Start();

                // Работа с элемнтами
                foreach (AROpeningHoleEntity arEntity in _arEntitiesToCreate)
                {
                    arEntity.CreateIntersectFamInstAndSetRevitParamsData(doc, arEntity.AR_OHE_HostElement);
                }

                // Работа с видом и выделением эл-в
                view = ViewZoomCreator.SpecialViewCreator(
                    app, 
                    _arEntitiesToCreate.Select(ent => ent.AR_OHE_HostElement).ToHashSet(new ElementComparerById()),
                    true);
                
                app.ActiveUIDocument.Selection.SetElementIds(
                    _arEntitiesToCreate
                    .Where(ent => doc.GetElement(ent.OHE_Element.Id).IsValidObject)
                    .Select(ent => ent.OHE_Element.Id)
                    .ToArray());

                trans.Commit();
            }

            ViewZoomCreator.SpecialViewOpener(app, view);

            return Result.Succeeded;
        }
    }
}

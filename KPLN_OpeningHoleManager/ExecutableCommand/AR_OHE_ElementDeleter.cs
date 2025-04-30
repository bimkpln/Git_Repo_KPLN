using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

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

        private readonly AROpeningHoleEntity[] _unionEntities;
        private List<ElementId> _arOpeningElems;

        public AR_OHE_ElementDeleter(AROpeningHoleEntity[] unionEntities)
        {
            _unionEntities = unionEntities;
        }

        public AR_OHE_ElementDeleter(List<ElementId> arOpeningElems)
        {
            _arOpeningElems = arOpeningElems;
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null) return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return Result.Cancelled;

            // Проверка на наличие объединенного отверстия. Если оно есть - значит ищем и удаляем ВСЕ отверстия, которые полностью 
            // поглащены объемом объединённого
            if (_unionEntities.Any())
            {
                _arOpeningElems = new List<ElementId>();
                foreach (AROpeningHoleEntity unionEntity in _unionEntities)
                {
                    Element[] hostElemOpeningColl = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Where(el => el is FamilyInstance fi 
                            && (fi.Symbol.FamilyName.StartsWith("199_Отвер") || fi.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие"))
                            && fi.Host.Id.IntegerValue == unionEntity.AR_OHE_HostElement.Id.IntegerValue)
                        .ToArray();

                    foreach(Element openingElem in hostElemOpeningColl)
                    {
                        if (openingElem.Id.IntegerValue == unionEntity.OHE_Element.Id.IntegerValue)
                            continue;
                    
                        Solid openingElemSolid = GeometryWorker.GetElemSolid(openingElem);
                        if (openingElemSolid == null)
                            continue;

                        Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(openingElemSolid, unionEntity.OHE_Solid, BooleanOperationsType.Intersect);
                        if (intersectSolid != null 
                            && intersectSolid.Volume > 0
                            && Math.Round(intersectSolid.Volume, 3) - Math.Round(openingElemSolid.Volume, 3) <= 0.01)
                            _arOpeningElems.Add(openingElem.Id);
                    }
                }
            }


            // Удаление элементов из списка на удаление
            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                foreach (ElementId elemId in _arOpeningElems)
                {
                    doc.Delete(elemId);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
